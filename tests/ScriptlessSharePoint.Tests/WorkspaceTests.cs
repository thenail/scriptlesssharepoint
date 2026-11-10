using ScriptlessSharePoint.Auth;
using ScriptlessSharePoint.Connections;

namespace ScriptlessSharePoint.Tests;

public sealed class WorkspaceTests : IDisposable
{
    private readonly string dir = Path.Combine(Path.GetTempPath(), "ssp-workspace-" + Guid.NewGuid().ToString("n"));
    private readonly ConnectionStore store;
    private readonly Workspace workspace;

    public WorkspaceTests()
    {
        store = new ConnectionStore(dir);
        workspace = new Workspace(store, () => IntPtr.Zero);
    }

    public void Dispose()
    {
        if (Directory.Exists(dir))
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    private static ConnectionDraft TokenDraft(string? id = null) => new()
    {
        Id = id,
        Name = "Contoso token",
        Kind = ConnectionKind.AccessToken,
        TenantUrl = "contoso",
    };

    [Fact]
    public void Saves_a_connection_with_its_token_encrypted()
    {
        var saved = workspace.Save(TokenDraft(), "Bearer aaa.bbb.ccc");
        Assert.Equal("https://contoso.sharepoint.com", saved.TenantUrl);
        Assert.Equal(saved.Id, store.Load().ActiveId);
        Assert.True(ConnectionDraft.From(saved).HasToken);
        Assert.DoesNotContain("aaa.bbb.ccc", File.ReadAllText(store.FilePath));
        Assert.Equal("aaa.bbb.ccc", Secrets.Unprotect(store.Load().Active!.ProtectedAccessToken));
    }

    [Fact]
    public void Editing_keeps_the_saved_token_unless_a_new_one_is_given()
    {
        var first = workspace.Save(TokenDraft(), "aaa.bbb.ccc").Id;
        var edit = TokenDraft(first);
        edit.Name = "Renamed";
        workspace.Save(edit, null);
        var profile = store.Load().Find(first)!;
        Assert.Equal("Renamed", profile.Name);
        Assert.Equal("aaa.bbb.ccc", Secrets.Unprotect(profile.ProtectedAccessToken));
        Assert.Single(store.Load().Connections);
    }

    [Fact]
    public void Changing_the_kind_forgets_the_account_and_token()
    {
        var first = workspace.Save(TokenDraft(), "aaa.bbb.ccc").Id;
        store.Update(d => d.Find(first)!.AccountName = "ada@contoso.com");
        workspace.Save(new ConnectionDraft
        {
            Id = first,
            Name = "Contoso app",
            Kind = ConnectionKind.Certificate,
            TenantUrl = "https://contoso.sharepoint.com",
            ClientId = Guid.NewGuid().ToString(),
            TenantId = "contoso.onmicrosoft.com",
            CertificateThumbprint = "ab cd ef",
        }, null);
        var profile = store.Load().Find(first)!;
        Assert.Equal(ConnectionKind.Certificate, profile.Kind);
        Assert.Null(profile.ProtectedAccessToken);
        Assert.Null(profile.AccountName);
        Assert.Equal("ABCDEF", profile.CertificateThumbprint);
    }

    [Fact]
    public void Invalid_connections_are_refused_with_what_is_missing()
    {
        var e = Assert.Throws<ConnectionException>(() =>
            workspace.Save(new ConnectionDraft { Name = "X", Kind = ConnectionKind.Certificate, TenantUrl = "contoso" }, null));
        Assert.Contains("app (client) ID", e.Message);
        Assert.Empty(store.Load().Connections);
    }

    [Fact]
    public void Deletes_a_connection_and_picks_another_active_one()
    {
        var a = workspace.Save(TokenDraft(), "aaa.bbb.ccc").Id;
        var b = workspace.Save(TokenDraft(), "aaa.bbb.ccc").Id;
        Assert.Equal(b, store.Load().ActiveId);
        var data = workspace.Delete(b);
        Assert.Equal(a, data.ActiveId);
        Assert.Single(store.Load().Connections);
    }

    [Fact]
    public async Task Sharepoint_calls_need_a_connection()
    {
        Assert.Equal("Add a connection first.", Assert.Throws<ConnectionException>(() => workspace.Current()).Message);
        await Assert.ThrowsAsync<ConnectionException>(() => workspace.StatusAsync());
    }

    [Fact]
    public async Task An_expired_token_says_to_sign_in_again()
    {
        var exp = DateTimeOffset.UtcNow.AddMinutes(-5).ToUnixTimeSeconds();
        var payload = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes($$"""{"aud":"https://contoso.sharepoint.com","exp":{{exp}}}""")).TrimEnd('=');
        workspace.Save(TokenDraft(), $"e30.{payload}.sig");

        Assert.Null(await workspace.StatusAsync());
        var e = await Assert.ThrowsAsync<SignInRequiredException>(() => workspace.Current().Client.GetSiteAsync("/sites/hr"));
        Assert.Contains("expired", e.Message);
    }

    [Fact]
    public void A_new_session_follows_a_changed_connection()
    {
        var id = workspace.Save(TokenDraft(), "aaa.bbb.ccc").Id;
        var first = workspace.Current();
        Assert.Same(first, workspace.Current());
        store.Update(d => d.Find(id)!.Name = "Changed elsewhere");
        Assert.NotSame(first, workspace.Current());
    }
}
