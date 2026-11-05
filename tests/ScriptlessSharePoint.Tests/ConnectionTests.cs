using System.Text;

using ScriptlessSharePoint.Auth;
using ScriptlessSharePoint.Connections;

namespace ScriptlessSharePoint.Tests;

public sealed class ConnectionTests : IDisposable
{
    private readonly string dir = Path.Combine(Path.GetTempPath(), "ssp-tests-" + Guid.NewGuid().ToString("n"));

    public void Dispose()
    {
        if (Directory.Exists(dir))
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void Store_round_trips_connections_and_the_active_one()
    {
        var store = new ConnectionStore(dir);
        Assert.Empty(store.Load().Connections);
        var a = new ConnectionProfile { Name = "A", TenantUrl = "https://a.sharepoint.com", ClientId = Guid.NewGuid().ToString() };
        var b = new ConnectionProfile { Name = "B", Kind = ConnectionKind.AccessToken, TenantUrl = "https://b.sharepoint.com", ProtectedAccessToken = Secrets.Protect("t") };
        store.Update(d =>
        {
            d.Connections.Add(a);
            d.Connections.Add(b);
            d.ActiveId = b.Id;
        });

        var loaded = new ConnectionStore(dir).Load();
        Assert.Equal(["A", "B"], loaded.Connections.Select(c => c.Name));
        Assert.Equal("B", loaded.Active?.Name);
        Assert.Equal("A", loaded.Find("a")?.Name);
        Assert.Equal("A", loaded.Find(a.Id)?.Name);
        Assert.Null(loaded.Find("nope"));
        Assert.Equal(ConnectionKind.AccessToken, loaded.Find("B")!.Kind);
        Assert.Contains("\"kind\": \"AccessToken\"", File.ReadAllText(store.FilePath));
    }

    [Fact]
    public void A_broken_file_loads_as_no_connections()
    {
        Directory.CreateDirectory(dir);
        File.WriteAllText(Path.Combine(dir, "connections.json"), "{ not json");
        Assert.Empty(new ConnectionStore(dir).Load().Connections);
    }

    [Fact]
    public void Secrets_round_trip_and_bad_input_reads_as_null()
    {
        var p = Secrets.Protect("secret value");
        Assert.NotEqual("secret value", p);
        Assert.Equal("secret value", Secrets.Unprotect(p));
        Assert.Null(Secrets.Unprotect("not base64!"));
        Assert.Null(Secrets.Unprotect(Convert.ToBase64String([1, 2, 3])));
        Assert.Null(Secrets.Unprotect(null));
    }

    [Fact]
    public void Validation_says_what_is_missing()
    {
        var p = new ConnectionProfile { Name = "X", TenantUrl = "contoso", Kind = ConnectionKind.Certificate, ClientId = Guid.NewGuid().ToString() };
        Assert.Contains("tenant ID", p.Validate());
        p.TenantId = "contoso.onmicrosoft.com";
        Assert.Contains("certificate", p.Validate());
        p.CertificateThumbprint = "AB";
        Assert.Null(p.Validate());

        var user = new ConnectionProfile { Name = "U", TenantUrl = "contoso", ClientId = Guid.NewGuid().ToString(), UseWindowsAccount = true, UseBrowser = true };
        Assert.Contains("browser", user.Validate());
        Assert.Contains("tenant", new ConnectionProfile { Name = "N", TenantUrl = "" }.Validate());
    }

    [Fact]
    public void Token_info_reads_claims()
    {
        var exp = DateTimeOffset.UtcNow.AddHours(1).ToUnixTimeSeconds();
        var token = Jwt($$"""{"aud":"https://contoso.sharepoint.com","exp":{{exp}},"upn":"ada@contoso.com","tid":"t1","scp":"AllSites.FullControl"}""");
        var info = TokenInfo.TryParse(token)!;
        Assert.Equal("https://contoso.sharepoint.com", info.Audience);
        Assert.Equal(exp, info.Expires!.Value.ToUnixTimeSeconds());
        Assert.Equal("ada@contoso.com", info.Subject);
        Assert.Equal("AllSites.FullControl", info.Permissions);
        Assert.True(info.IsForSharePoint);

        Assert.False(TokenInfo.TryParse(Jwt("""{"aud":"https://graph.microsoft.com","roles":["Sites.Read.All"]}"""))!.IsForSharePoint);
        Assert.Null(TokenInfo.TryParse("nope"));
        Assert.Null(TokenInfo.TryParse("a.%%%.c"));
    }

    [Fact]
    public async Task An_expired_pasted_token_asks_for_a_new_one()
    {
        var token = Jwt($$"""{"aud":"https://contoso.sharepoint.com","exp":{{DateTimeOffset.UtcNow.AddMinutes(-5).ToUnixTimeSeconds()}}}""");
        var p = new ConnectionProfile { Name = "T", Kind = ConnectionKind.AccessToken, TenantUrl = "contoso", ProtectedAccessToken = Secrets.Protect(token) };
        var e = await Assert.ThrowsAsync<SignInRequiredException>(() => new AccessTokenProvider(p).GetTokenAsync("https://contoso.sharepoint.com", false));
        Assert.Contains("expired", e.Message);

        var fresh = Jwt($$"""{"aud":"https://contoso.sharepoint.com","exp":{{DateTimeOffset.UtcNow.AddMinutes(30).ToUnixTimeSeconds()}}}""");
        p.ProtectedAccessToken = Secrets.Protect(fresh);
        Assert.Equal(fresh, await new AccessTokenProvider(p).GetTokenAsync("https://contoso.sharepoint.com", false));
    }

    [Fact]
    public void Thumbprints_are_normalized() =>
        Assert.Equal("AB12CD", Certificates.Normalize(" ab:12 cd‎"));

    private static string Jwt(string payload)
    {
        static string B64(string s) => Convert.ToBase64String(Encoding.UTF8.GetBytes(s)).TrimEnd('=').Replace('+', '-').Replace('/', '_');
        return $"{B64("""{"alg":"none"}""")}.{B64(payload)}.sig";
    }
}
