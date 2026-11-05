using System.Text.Json.Nodes;

using ScriptlessSharePoint.Bridge;
using ScriptlessSharePoint.Connections;

namespace ScriptlessSharePoint.Tests;

public sealed class AppBridgeTests : IDisposable
{
    private sealed class FakeHost : IHostUi
    {
        public List<string> Opened { get; } = [];
        public IntPtr WindowHandle => IntPtr.Zero;
        public string? StartSite => "https://contoso.sharepoint.com/sites/hr";
        public string McpCommand => "scriptless-sharepoint-mcp";
        public string? PickFile(string title, string filter) => null;
        public string? PickSaveFile(string title, string filter, string fileName) => null;
        public void OpenUrl(string url) => Opened.Add(url);
        public void Notify(JsonObject message) { }
    }

    private readonly string dir = Path.Combine(Path.GetTempPath(), "ssp-bridge-" + Guid.NewGuid().ToString("n"));
    private readonly FakeHost host = new();
    private readonly ConnectionStore store;
    private readonly AppBridge bridge;
    private int id;

    public AppBridgeTests()
    {
        store = new ConnectionStore(dir);
        bridge = new AppBridge(store, host);
    }

    public void Dispose()
    {
        if (Directory.Exists(dir))
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    private async Task<JsonObject> Call(string method, JsonObject? p = null)
    {
        var request = new JsonObject { ["jsonrpc"] = "2.0", ["id"] = ++id, ["method"] = method, ["params"] = p ?? [] };
        var reply = JsonNode.Parse((await bridge.HandleAsync(request.ToJsonString()))!)!.AsObject();
        Assert.Equal(id, reply["id"]!.GetValue<int>());
        return reply;
    }

    private async Task<JsonNode?> Result(string method, JsonObject? p = null)
    {
        var reply = await Call(method, p);
        Assert.True(reply["error"] is null, reply["error"]?.ToJsonString());
        return reply["result"];
    }

    private async Task<string> ErrorMessage(string method, JsonObject? p = null) =>
        (await Call(method, p))["error"]!["message"]!.GetValue<string>();

    private static JsonObject TokenConnection(string? id = null) => new()
    {
        ["connection"] = new JsonObject
        {
            ["id"] = id,
            ["name"] = "Contoso token",
            ["kind"] = "AccessToken",
            ["tenantUrl"] = "contoso",
        },
        ["token"] = "Bearer aaa.bbb.ccc",
    };

    [Fact]
    public async Task Info_says_how_to_start_the_mcp_server()
    {
        var info = (await Result("app.info"))!;
        Assert.Equal("Scriptless SharePoint", info["name"]!.GetValue<string>());
        Assert.Equal("scriptless-sharepoint-mcp", info["mcpCommand"]!.GetValue<string>());
        Assert.Equal("https://contoso.sharepoint.com/sites/hr", info["startSite"]!.GetValue<string>());
    }

    [Fact]
    public async Task Saves_a_connection_with_its_token_encrypted_and_never_sent_back()
    {
        var saved = (await Result("connections.save", TokenConnection()))!;
        var connection = saved["connection"]!;
        Assert.Equal("https://contoso.sharepoint.com", connection["tenantUrl"]!.GetValue<string>());
        Assert.True(connection["hasToken"]!.GetValue<bool>());
        Assert.Equal(connection["id"]!.GetValue<string>(), saved["activeId"]!.GetValue<string>());

        var list = (await Result("connections.list"))!;
        Assert.DoesNotContain("aaa.bbb.ccc", list.ToJsonString());
        Assert.DoesNotContain("aaa.bbb.ccc", File.ReadAllText(store.FilePath));
        Assert.Equal("aaa.bbb.ccc", Secrets.Unprotect(store.Load().Active!.ProtectedAccessToken));
    }

    [Fact]
    public async Task Editing_keeps_the_saved_token_unless_a_new_one_is_given()
    {
        var first = (await Result("connections.save", TokenConnection()))!["connection"]!["id"]!.GetValue<string>();
        var edit = TokenConnection(first);
        edit.Remove("token");
        edit["connection"]!["name"] = "Renamed";
        await Result("connections.save", edit);
        var profile = store.Load().Find(first)!;
        Assert.Equal("Renamed", profile.Name);
        Assert.Equal("aaa.bbb.ccc", Secrets.Unprotect(profile.ProtectedAccessToken));
        Assert.Single(store.Load().Connections);
    }

    [Fact]
    public async Task Changing_the_kind_forgets_the_account_and_token()
    {
        var first = (await Result("connections.save", TokenConnection()))!["connection"]!["id"]!.GetValue<string>();
        store.Update(d => d.Find(first)!.AccountName = "ada@contoso.com");
        await Result("connections.save", new JsonObject
        {
            ["connection"] = new JsonObject
            {
                ["id"] = first,
                ["name"] = "Contoso app",
                ["kind"] = "Certificate",
                ["tenantUrl"] = "https://contoso.sharepoint.com",
                ["clientId"] = Guid.NewGuid().ToString(),
                ["tenantId"] = "contoso.onmicrosoft.com",
                ["certificateThumbprint"] = "ab cd ef",
            },
        });
        var profile = store.Load().Find(first)!;
        Assert.Equal(ConnectionKind.Certificate, profile.Kind);
        Assert.Null(profile.ProtectedAccessToken);
        Assert.Null(profile.AccountName);
        Assert.Equal("ABCDEF", profile.CertificateThumbprint);
    }

    [Fact]
    public async Task Invalid_connections_are_refused_with_what_is_missing()
    {
        var message = await ErrorMessage("connections.save", new JsonObject
        {
            ["connection"] = new JsonObject { ["name"] = "X", ["kind"] = "Certificate", ["tenantUrl"] = "contoso" },
        });
        Assert.Contains("app (client) ID", message);
        Assert.Empty(store.Load().Connections);
    }

    [Fact]
    public async Task Deletes_a_connection_and_picks_another_active_one()
    {
        var a = (await Result("connections.save", TokenConnection()))!["connection"]!["id"]!.GetValue<string>();
        var b = (await Result("connections.save", TokenConnection()))!["connection"]!["id"]!.GetValue<string>();
        Assert.Equal(b, store.Load().ActiveId);
        var list = (await Result("connections.delete", new JsonObject { ["id"] = b }))!;
        Assert.Equal(a, list["activeId"]!.GetValue<string>());
        Assert.Single(store.Load().Connections);
    }

    [Fact]
    public async Task Sharepoint_calls_need_a_connection()
    {
        Assert.Equal("Add a connection first.", await ErrorMessage("sites.search", new JsonObject { ["query"] = "hr" }));
        Assert.Equal("Add a connection first.", await ErrorMessage("session.status"));
    }

    [Fact]
    public async Task An_expired_token_says_to_sign_in_again()
    {
        var exp = DateTimeOffset.UtcNow.AddMinutes(-5).ToUnixTimeSeconds();
        var payload = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes($$"""{"aud":"https://contoso.sharepoint.com","exp":{{exp}}}""")).TrimEnd('=');
        var p = TokenConnection();
        p["token"] = $"e30.{payload}.sig";
        await Result("connections.save", p);

        var reply = await Call("sites.get", new JsonObject { ["url"] = "/sites/hr" });
        Assert.Contains("expired", reply["error"]!["message"]!.GetValue<string>());
        Assert.True(reply["error"]!["data"]!["signInRequired"]!.GetValue<bool>());
    }

    [Fact]
    public async Task Inspects_tokens_and_tenants()
    {
        Assert.Null(await Result("tokens.inspect", new JsonObject { ["token"] = "nope" }));
        var tenant = (await Result("tenant.normalize", new JsonObject { ["input"] = "https://contoso-admin.sharepoint.com/_layouts" }))!;
        Assert.Equal("https://contoso.sharepoint.com", tenant["tenantUrl"]!.GetValue<string>());
        Assert.Equal("https://contoso-admin.sharepoint.com", tenant["adminUrl"]!.GetValue<string>());
    }

    [Fact]
    public async Task Opens_only_https_links()
    {
        Assert.Contains("https://", await ErrorMessage("shell.openUrl", new JsonObject { ["url"] = "file:///C:/Windows/notepad.exe" }));
        await Result("shell.openUrl", new JsonObject { ["url"] = "https://contoso.sharepoint.com/sites/hr" });
        Assert.Equal(["https://contoso.sharepoint.com/sites/hr"], host.Opened);
    }

    [Fact]
    public async Task Protocol_errors()
    {
        Assert.Contains("Invalid request", await bridge.HandleAsync("{ nope"));
        Assert.Contains("Method not found", await ErrorMessage("sites.delete_everything"));
        // notifications get no reply
        Assert.Null(await bridge.HandleAsync("""{"jsonrpc":"2.0","method":"app.info"}"""));
    }
}
