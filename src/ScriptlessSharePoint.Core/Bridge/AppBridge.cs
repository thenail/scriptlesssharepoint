// The web UI's backend: JSON-RPC 2.0 requests from the page (WebView2
// postMessage), one method per operation. The same operations as the MCP
// server's tools, plus connections, sign-in and the host's dialogs.
// Requests are handled on the UI thread, so sign-in dialogs open there.
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

using ScriptlessSharePoint.Auth;
using ScriptlessSharePoint.Connections;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Bridge;

/// <summary>A failure the UI shows as it is.</summary>
public sealed class BridgeException(string message) : Exception(message);

public sealed class AppBridge
{
    public static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() },
    };

    private const int MethodNotFound = -32601;
    private const int InvalidRequest = -32600;
    private const int Failed = -32000;

    private readonly ConnectionStore store;
    private readonly IHostUi host;
    private readonly Dictionary<string, Func<JsonObject, Task<object?>>> methods;
    /// <summary>Accounts signed in with by connections.test, applied when that connection is saved.</summary>
    private readonly Dictionary<string, (string? Id, string? Name)> testedAccounts = [];
    private Session? session;

    public AppBridge(ConnectionStore store, IHostUi host)
    {
        this.store = store;
        this.host = host;
        methods = new()
        {
            ["app.info"] = _ => Task.FromResult<object?>(new
            {
                name = AppInfo.Name,
                version = AppInfo.Version,
                builtInClientId = AppInfo.BuiltInClientId,
                dataDir = AppInfo.DataDir,
                mcpCommand = host.McpCommand,
                startSite = host.StartSite,
            }),
            ["connections.list"] = _ => Task.FromResult<object?>(ListConnections()),
            ["connections.save"] = p => Task.FromResult<object?>(Save(p)),
            ["connections.delete"] = DeleteAsync,
            ["connections.setActive"] = p => Task.FromResult<object?>(SetActive(Required(p, "id"))),
            ["connections.test"] = TestAsync,
            ["session.status"] = StatusAsync,
            ["session.signIn"] = async _ =>
            {
                var s = Current();
                s.Identity = await s.Client.WhoAmIAsync();
                await PrepareAdminTokenAsync(s.Profile, s.Tokens);
                return new { identity = s.Identity, accountName = s.Profile.AccountName };
            },
            ["session.signOut"] = async _ =>
            {
                var s = Current();
                await s.Tokens.SignOutAsync();
                s.Identity = null;
                return null;
            },
            ["sites.search"] = async p =>
            {
                var c = Current().Client;
                var query = Optional(p, "query");
                var cursor = Optional(p, "cursor");
                var limit = Int(p, "limit", 50);
                return Optional(p, "scope") == "admin"
                    ? await c.ListSitesAsAdminAsync(query, limit, cursor)
                    : await c.SearchSitesAsync(query, limit, cursor);
            },
            ["sites.get"] = async p => await Current().Client.GetSiteAsync(Required(p, "url")),
            ["catalogs.list"] = async _ => await Current().Client.ListSiteAppCatalogsAsync(),
            ["catalogs.enable"] = async p => await Current().Client.EnableSiteAppCatalogAsync(Required(p, "url")),
            ["catalogs.disable"] = async p => await Current().Client.DisableSiteAppCatalogAsync(Required(p, "url")),
            ["certificates.pickPfx"] = _ =>
            {
                var path = host.PickFile("Import a certificate", "Certificates with a private key (*.pfx;*.p12)|*.pfx;*.p12");
                return Task.FromResult<object?>(path is null ? null : new { path, fileName = Path.GetFileName(path) });
            },
            ["certificates.importPfx"] = p =>
            {
                var thumbprint = Certificates.ImportPfx(Required(p, "path"), Optional(p, "password"));
                return Task.FromResult<object?>(new { thumbprint, description = Certificates.Describe(thumbprint) });
            },
            ["certificates.create"] = p =>
            {
                var name = Optional(p, "name") ?? AppInfo.Name;
                var path = host.PickSaveFile(
                    "Save the certificate's public key (.cer) to upload to the app registration",
                    "Certificate (*.cer)|*.cer",
                    name.Replace(' ', '-') + ".cer");
                if (path is null)
                {
                    return Task.FromResult<object?>(null);
                }
                var thumbprint = Certificates.CreateSelfSigned(name, path);
                return Task.FromResult<object?>(new { thumbprint, cerPath = path, description = Certificates.Describe(thumbprint) });
            },
            ["certificates.describe"] = p =>
            {
                var tp = Optional(p, "thumbprint");
                bool ok;
                try
                {
                    ok = tp is not null && Certificates.Find(tp) is not null;
                }
                catch (InvalidOperationException)
                {
                    ok = false;
                }
                return Task.FromResult<object?>(new { ok, description = Certificates.Describe(tp) });
            },
            ["tokens.inspect"] = p => Task.FromResult<object?>(
                TokenInfo.TryParse(ConnectionDto.CleanToken(Optional(p, "token"))) is { } info
                    ? new { info.Audience, info.Expires, info.Subject, info.TenantId, info.Permissions, info.IsForSharePoint, expired = info.Expires < DateTimeOffset.UtcNow }
                    : null),
            ["tenant.normalize"] = p => Task.FromResult<object?>(
                TenantUrls.TryNormalizeTenantUrl(Optional(p, "input"), out var url)
                    ? new { tenantUrl = url, adminUrl = TenantUrls.AdminUrl(url) }
                    : null),
            ["shell.openUrl"] = p =>
            {
                var url = Required(p, "url");
                if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeHttps)
                {
                    throw new BridgeException("Only https:// links open in the browser.");
                }
                host.OpenUrl(uri.AbsoluteUri);
                return Task.FromResult<object?>(null);
            },
        };
    }

    /// <summary>Handles one message from the page; returns the response to post back, or null.</summary>
    public async Task<string?> HandleAsync(string? json)
    {
        JsonObject? msg;
        try
        {
            msg = JsonNode.Parse(json ?? "") as JsonObject;
        }
        catch (JsonException)
        {
            msg = null;
        }
        var id = msg?["id"]?.DeepClone();
        var method = msg?["method"]?.GetValueKind() == JsonValueKind.String ? msg["method"]!.GetValue<string>() : null;
        if (msg is null || method is null)
        {
            return Error(id, InvalidRequest, "Invalid request");
        }
        if (!methods.TryGetValue(method, out var run))
        {
            return id is null ? null : Error(id, MethodNotFound, $"Method not found: {method}");
        }
        var started = Environment.TickCount64;
        try
        {
            var result = await run(msg["params"] as JsonObject ?? []);
            Diagnostics.Log.Write("bridge", $"{method}: ok in {Environment.TickCount64 - started} ms");
            return id is null
                ? null
                : new JsonObject { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = JsonSerializer.SerializeToNode(result, Json) }.ToJsonString();
        }
        catch (Exception e)
        {
            var message = e is BridgeException ? e.Message : SharePointErrors.Describe(e);
            var signIn = e is SignInRequiredException || e.InnerException is SignInRequiredException;
            Diagnostics.Log.Write("bridge", $"{method}: failed in {Environment.TickCount64 - started} ms: {e.GetType().Name}: {message}");
            return id is null ? null : Error(id, Failed, message, signIn);
        }
    }

    public IReadOnlyList<string> Methods => [.. methods.Keys];

    // ---- connections ----

    private object ListConnections()
    {
        var data = store.Load();
        return new { activeId = data.Active?.Id, connections = data.Connections.Select(ConnectionDto.From) };
    }

    private object Save(JsonObject p)
    {
        var dto = Dto(p);
        ConnectionProfile? saved = null;
        store.Update(data =>
        {
            var existing = data.Connections.FirstOrDefault(c => c.Id == dto.Id);
            var profile = dto.ToProfile(existing, Optional(p, "token"));
            if (testedAccounts.Remove(profile.Id, out var account) && profile.Kind == ConnectionKind.User)
            {
                (profile.AccountId, profile.AccountName) = account;
            }
            if (profile.Validate() is { } problem)
            {
                throw new BridgeException(problem);
            }
            if (existing is null)
            {
                data.Connections.Add(profile);
            }
            else
            {
                data.Connections[data.Connections.IndexOf(existing)] = profile;
            }
            data.ActiveId = profile.Id;
            saved = profile;
        });
        session = null; // the next call starts a session with the saved connection
        return new { connection = ConnectionDto.From(saved!), activeId = saved!.Id };
    }

    private Task<object?> DeleteAsync(JsonObject p)
    {
        var id = Required(p, "id");
        var profile = store.Load().Find(id) ?? throw new BridgeException("That connection doesn't exist any more.");
        // the account stays signed in: other connections (and AI apps) may use
        // it, and Windows would refuse it silently until signed in again
        store.Update(d =>
        {
            d.Connections.RemoveAll(c => c.Id == profile.Id);
            if (d.ActiveId == profile.Id)
            {
                d.ActiveId = d.Connections.FirstOrDefault()?.Id;
            }
        });
        if (session?.Profile.Id == profile.Id)
        {
            session = null;
        }
        return Task.FromResult<object?>(ListConnections());
    }

    private object SetActive(string id)
    {
        var data = store.Update(d =>
        {
            if (d.Find(id) is null)
            {
                throw new BridgeException("That connection doesn't exist any more.");
            }
            d.ActiveId = id;
        });
        return new { activeId = data.ActiveId };
    }

    /// <summary>Signs in with a connection as edited (not yet saved) and says who it is.</summary>
    private async Task<object?> TestAsync(JsonObject p)
    {
        var dto = Dto(p);
        var saved = dto.Id is null ? null : store.Load().Connections.FirstOrDefault(c => c.Id == dto.Id);
        var profile = dto.ToProfile(saved, Optional(p, "token"));
        if (profile.Validate() is { } problem)
        {
            throw new BridgeException(problem);
        }
        var tokens = TokenProviders.Create(profile, () => host.WindowHandle);
        var identity = await new SharePointClient(profile, tokens, interactive: true).WhoAmIAsync();
        await PrepareAdminTokenAsync(profile, tokens);
        testedAccounts[profile.Id] = (profile.AccountId, profile.AccountName);
        return new { id = profile.Id, identity, accountName = profile.AccountName };
    }

    // ---- the session ----

    /// <summary>The active connection's session: made on first use, again when the connection changed.</summary>
    private Session Current()
    {
        var data = store.Load();
        var profile = data.Active ?? throw new BridgeException("Add a connection first.");
        if (profile.Validate() is { } problem)
        {
            throw new BridgeException($"The connection \"{profile.Name}\" isn't complete: {problem}");
        }
        var fingerprint = JsonSerializer.Serialize(profile);
        if (session is null || session.Profile.Id != profile.Id || session.Fingerprint != fingerprint)
        {
            session = new Session(profile, fingerprint, () => host.WindowHandle, OnAccountChanged);
        }
        return session;
    }

    /// <summary>
    /// The admin centre is its own resource (https://contoso-admin.sharepoint.com):
    /// get its token while the user is signing in, so a prompt for it, if Windows
    /// needs one, comes now rather than at the first admin operation. Its failure
    /// isn't the sign-in's: the admin operations report it when they're used.
    /// </summary>
    private static async Task PrepareAdminTokenAsync(ConnectionProfile profile, ITokenProvider tokens)
    {
        if (profile.Kind != ConnectionKind.User)
        {
            return;
        }
        try
        {
            await tokens.GetTokenAsync(TenantUrls.Resource(profile.EffectiveAdminUrl), interactive: true);
        }
        catch (Exception e)
        {
            Diagnostics.Log.Write("auth", $"admin centre token after sign-in: {e.GetType().Name}: {SharePointErrors.Describe(e)}");
        }
    }

    private async Task<object?> StatusAsync(JsonObject p)
    {
        var s = Current();
        if (s.Identity is null)
        {
            try
            {
                s.Identity = await s.Silent.WhoAmIAsync();
            }
            catch (Exception)
            {
                // not signed in (or offline): the UI offers Sign in
            }
        }
        return new { connectionId = s.Profile.Id, identity = s.Identity, accountName = s.Profile.AccountName };
    }

    private void OnAccountChanged(ConnectionProfile profile)
    {
        store.Update(d =>
        {
            var saved = d.Connections.FirstOrDefault(c => c.Id == profile.Id);
            if (saved is not null)
            {
                saved.AccountId = profile.AccountId;
                saved.AccountName = profile.AccountName;
            }
        });
        if (session?.Profile.Id == profile.Id)
        {
            // the account is saved now; the session stays the one signed in
            session.Fingerprint = JsonSerializer.Serialize(store.Load().Find(profile.Id));
        }
        host.Notify(new JsonObject
        {
            ["jsonrpc"] = "2.0",
            ["method"] = "session.accountChanged",
            ["params"] = new JsonObject { ["connectionId"] = profile.Id, ["accountName"] = profile.AccountName },
        });
    }

    private sealed class Session(ConnectionProfile profile, string fingerprint, Func<IntPtr> window, Action<ConnectionProfile> accountChanged)
    {
        public ConnectionProfile Profile { get; } = profile;
        public string Fingerprint { get; set; } = fingerprint;
        public ITokenProvider Tokens { get; } = TokenProviders.Create(profile, window, accountChanged);
        public Identity? Identity { get; set; }
        public SharePointClient Client => client ??= new SharePointClient(Profile, Tokens, interactive: true);
        public SharePointClient Silent => silent ??= new SharePointClient(Profile, Tokens, interactive: false);
        private SharePointClient? client;
        private SharePointClient? silent;
    }

    // ---- arguments ----

    private static ConnectionDto Dto(JsonObject p) =>
        p["connection"]?.Deserialize<ConnectionDto>(Json) ?? throw new BridgeException("connection: required");

    private static string? Optional(JsonObject p, string name) =>
        p[name]?.GetValueKind() == JsonValueKind.String && p[name]!.GetValue<string>().Trim() is { Length: > 0 } s ? s : null;

    private static string Required(JsonObject p, string name) =>
        Optional(p, name) ?? throw new BridgeException($"{name}: required");

    private static int Int(JsonObject p, string name, int fallback) =>
        p[name]?.GetValueKind() == JsonValueKind.Number && p[name]!.AsValue().TryGetValue<int>(out var n) ? n : fallback;

    private static string Error(JsonNode? id, int code, string message, bool signIn = false)
    {
        var error = new JsonObject { ["code"] = code, ["message"] = message };
        if (signIn)
        {
            error["data"] = new JsonObject { ["signInRequired"] = true };
        }
        return new JsonObject { ["jsonrpc"] = "2.0", ["id"] = id, ["error"] = error }.ToJsonString();
    }
}
