// The server's tools: list and read connections and sites, enable or remove
// site collection app catalogs, and make plans (SharePointTools.Plans.cs). Calls run one at a time; failures come
// back as results with isError, never as protocol errors.
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Mcp;

/// <summary>MCP tool annotations (hints for the host's confirmations).</summary>
public sealed record ToolAnnotations(bool ReadOnlyHint, bool DestructiveHint, bool IdempotentHint, bool OpenWorldHint);

/// <summary>A tool as tools/list describes it.</summary>
public sealed record McpTool(string Name, string Title, string Description, JsonObject InputSchema, ToolAnnotations Annotations)
{
    public JsonObject ToJson() => new()
    {
        ["name"] = Name,
        ["title"] = Title,
        ["description"] = Description,
        ["inputSchema"] = InputSchema.DeepClone(),
        ["annotations"] = new JsonObject
        {
            ["title"] = Title,
            ["readOnlyHint"] = Annotations.ReadOnlyHint,
            ["destructiveHint"] = Annotations.DestructiveHint,
            ["idempotentHint"] = Annotations.IdempotentHint,
            ["openWorldHint"] = Annotations.OpenWorldHint,
        },
    };
}

/// <summary>A tools/call result: text content, maybe an error.</summary>
public sealed record ToolResult(string Text, bool IsError = false)
{
    public JsonObject ToJson() => new()
    {
        ["content"] = new JsonArray(new JsonObject { ["type"] = "text", ["text"] = Text }),
        ["isError"] = IsError,
    };
}

/// <summary>Bad arguments: the message says which and why.</summary>
public sealed class ToolArgumentException(string message) : Exception(message);

public sealed partial class SharePointTools
{
    private static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) },
    };

    private static readonly ToolAnnotations Read = new(true, false, true, true);

    private const string ConnectionDescription =
        "Name of a saved connection (see list_connections). Leave it out to use the active connection.";

    private const string UrlDescription =
        "URL of the site, e.g. https://contoso.sharepoint.com/sites/hr. A subsite, page or library URL in the site works too.";

    /// <summary>Every tool, in the order tools/list gives them.</summary>
    public static readonly IReadOnlyList<McpTool> All =
    [
        new("list_connections", "List connections",
            "Lists the SharePoint connections saved in the Scriptless SharePoint app: name, how it signs in, tenant, account, and which one is active (used when a tool is given no `connection`).",
            Schema(), Read with { OpenWorldHint = false }),
        new("get_connection", "Check a connection",
            "Signs in with a connection and says who it acts as (user or app) and the tenant's name. Use it to check that a connection works.",
            Schema(Connection()), Read),
        new("search_sites", "Search sites",
            "Finds site collections, a page at a time: pass a result's `next` as `cursor` for the next page (no `next`: that was the last). `scope` \"search\" (default) uses SharePoint search: any account, only the sites it can see, matching words that start with `query`'s words; new sites can take minutes to appear. `scope` \"admin\" searches every site in the tenant in the admin centre (SharePoint administrators only): sites whose title contains all of `query`'s words, then those whose URL does; without `query`, all sites by title. A word too common to search on is refused: use more of the name. OneDrive sites are left out.",
            Schema(
                ("query", Str("Words in the site's title or URL."), false),
                ("scope", new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("search", "admin"), ["default"] = "search" }, false),
                ("limit", Int(1, 200, 25), false),
                ("cursor", Str("The `next` of the previous page, for the page after it."), false),
                Connection()),
            Read),
        new("get_site", "Get site",
            "Gets a site collection (for any URL in it): its URL, title, template, ID, and whether it has its own app catalog (`appCatalog`: enabled or notEnabled; `appCatalogSource` \"tenant\" is authoritative, \"site\" means the connection isn't an admin and it was judged from the site's libraries).",
            Schema(("url", Str(UrlDescription), true), Connection()), Read),
        new("list_site_app_catalogs", "List site collection app catalogs",
            "Lists every site collection in the tenant that has its own app catalog. Needs a SharePoint administrator.",
            Schema(Connection()), Read),
        new("enable_site_app_catalog", "Enable site collection app catalog",
            "Gives a site collection its own app catalog (an Apps for SharePoint library), so SharePoint Framework solutions can be added to that site only, as Add-SPOSiteCollectionAppCatalog does. Needs a SharePoint administrator. Does nothing if the site already has one.",
            Schema(("url", Str(UrlDescription), true), Connection()),
            new(false, false, true, true)),
        new("disable_site_app_catalog", "Remove site collection app catalog",
            "Removes a site collection's app catalog (Remove-SPOSiteCollectionAppCatalog). Its Apps for SharePoint library and files stay in the site, but it stops being an app catalog, and solutions added from it may stop working on the site. Needs a SharePoint administrator. Only when the user asks for it.",
            Schema(("url", Str(UrlDescription), true), Connection()),
            new(false, true, true, true)),
        new("open_app", "Open Scriptless SharePoint",
            "Opens the Scriptless SharePoint app on this PC, where the user signs in and manages connections; with `url`, it shows that site; with `plan`, that plan, for the user to look over and run.",
            Schema(("url", Str("A site to show (optional)."), false), ("plan", Str("A plan's id to show (optional)."), false)),
            Read with { OpenWorldHint = false }),
        .. PlanTools(),
    ];

    private readonly IToolBackend backend;
    private readonly Func<string, Task<bool>> confirm;
    private readonly SemaphoreSlim queue = new(1, 1);

    public SharePointTools(IToolBackend backend, Func<string, Task<bool>> confirm)
    {
        this.backend = backend;
        this.confirm = confirm;
    }

    /// <summary>Who's calling, from initialize.</summary>
    public string Client { get; set; } = "agent";

    public static JsonArray ToolListJson() => new([.. All.Select(t => (JsonNode)t.ToJson())]);

    public bool Has(string name) => All.Any(t => t.Name == name);

    /// <summary>Runs a call; never throws.</summary>
    public async Task<ToolResult> CallAsync(string name, JsonObject args, CancellationToken ct = default)
    {
        await queue.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            return await RunAsync(name, new Args(name, args), ct).ConfigureAwait(false);
        }
        catch (ToolArgumentException e)
        {
            return new ToolResult(e.Message, true);
        }
        catch (Exception e)
        {
            return new ToolResult(SharePointErrors.Describe(e), true);
        }
        finally
        {
            queue.Release();
        }
    }

    private async Task<ToolResult> RunAsync(string name, Args a, CancellationToken ct)
    {
        switch (name)
        {
            case "list_connections":
            {
                var list = backend.ListConnections();
                if (list.Count == 0)
                {
                    return new ToolResult(
                        $"There are no connections yet. Ask the user to open {AppInfo.Name} (open_app) and add one: it signs in to SharePoint there.");
                }
                return Data(new { connections = list }, "connections");
            }
            case "get_connection":
            {
                var connection = a.String("connection");
                var view = backend.Resolve(connection);
                var identity = await backend.WhoAmIAsync(connection, ct).ConfigureAwait(false);
                return Data(new { connection = view, identity }, "connection");
            }
            case "search_sites":
            {
                var scope = a.Enum("scope", ["search", "admin"], "search");
                var result = await backend.SearchSitesAsync(
                    a.String("connection"),
                    a.String("query"),
                    a.Int("limit", 1, 200, 25),
                    a.String("cursor"),
                    scope == "admin",
                    ct).ConfigureAwait(false);
                return Data(new
                {
                    result.Sites,
                    result.Next,
                    result.Total,
                    totalIsEstimate = result.TotalIsEstimate ? true : (bool?)null,
                    result.Source,
                }, "sites");
            }
            case "get_site":
                return Data(await backend.GetSiteAsync(a.String("connection"), a.Required("url"), ct).ConfigureAwait(false), "site");
            case "list_site_app_catalogs":
            {
                var catalogs = await backend.ListSiteAppCatalogsAsync(a.String("connection"), ct).ConfigureAwait(false);
                return Data(new { siteAppCatalogs = catalogs, total = catalogs.Count }, "site_app_catalogs");
            }
            case "enable_site_app_catalog":
            {
                var change = await backend.EnableSiteAppCatalogAsync(a.String("connection"), a.Required("url"), ct).ConfigureAwait(false);
                return Data(change, "result");
            }
            case "disable_site_app_catalog":
            {
                var connection = a.String("connection");
                var url = a.Required("url");
                var view = backend.Resolve(connection);
                var allowed = await confirm(
                    $"Remove the site collection app catalog from {url} (connection \"{view.Name}\", {view.TenantUrl})? Its Apps for SharePoint library stays, but solutions added from it may stop working on the site.")
                    .ConfigureAwait(false);
                if (!allowed)
                {
                    return new ToolResult("The user declined: the app catalog was left as it is.", true);
                }
                var change = await backend.DisableSiteAppCatalogAsync(connection, url, ct).ConfigureAwait(false);
                return Data(change, "result");
            }
            case "open_app":
            {
                var url = a.String("url");
                var plan = a.String("plan");
                if (plan is not null && backend.Plans.Load(plan) is null)
                {
                    return new ToolResult($"There's no plan \"{plan}\" (list_plans lists them).", true);
                }
                await backend.OpenAppAsync(url, plan).ConfigureAwait(false);
                return new ToolResult(JsonSerializer.Serialize(new { opened = true, url, plan }, Json));
            }
            default:
                return await RunPlanToolAsync(name, a).ConfigureAwait(false) ?? new ToolResult($"Unknown tool {name}.", true);
        }
    }

    /// <summary>A result with values from SharePoint, marked as data (see the server's instructions).</summary>
    private static ToolResult Data(object value, string kind) =>
        new($"<sharepoint_data kind=\"{kind}\">\n{JsonSerializer.Serialize(value, Json)}\n</sharepoint_data>");

    private static JsonObject Schema(params (string Name, JsonObject Schema, bool Required)[] properties)
    {
        var props = new JsonObject();
        var required = new JsonArray();
        foreach (var (name, schema, isRequired) in properties)
        {
            props[name] = schema;
            if (isRequired)
            {
                required.Add(name);
            }
        }
        var result = new JsonObject { ["type"] = "object", ["properties"] = props, ["additionalProperties"] = false };
        if (required.Count > 0)
        {
            result["required"] = required;
        }
        return result;
    }

    private static (string, JsonObject, bool) Connection() => ("connection", Str(ConnectionDescription), false);

    private static JsonObject Str(string description) =>
        new() { ["type"] = "string", ["minLength"] = 1, ["description"] = description };

    private static JsonObject Int(int min, int max, int fallback, string? description = null)
    {
        var o = new JsonObject { ["type"] = "integer", ["minimum"] = min, ["maximum"] = max, ["default"] = fallback };
        if (description is not null)
        {
            o["description"] = description;
        }
        return o;
    }

    /// <summary>Reads and checks a call's arguments against its schema.</summary>
    private sealed class Args
    {
        private readonly string tool;
        private readonly JsonObject args;

        public Args(string tool, JsonObject args)
        {
            this.tool = tool;
            this.args = args;
            var known = All.First(t => t.Name == tool).InputSchema["properties"]!.AsObject();
            var unknown = args.Select(p => p.Key).Where(k => !known.ContainsKey(k)).ToList();
            if (unknown.Count > 0)
            {
                throw Invalid($"unknown argument{(unknown.Count == 1 ? "" : "s")} {string.Join(", ", unknown)}");
            }
        }

        public string? String(string name)
        {
            var node = args[name];
            if (node is null)
            {
                return null;
            }
            if (node.GetValueKind() != JsonValueKind.String)
            {
                throw Invalid($"{name}: expected a string");
            }
            var s = node.GetValue<string>().Trim();
            return s.Length == 0 ? null : s;
        }

        public string Required(string name) => String(name) ?? throw Invalid($"{name}: required");

        /// <summary>An argument as it came (arrays and objects), or null.</summary>
        public JsonNode? Node(string name) => args[name];

        public int Int(string name, int min, int max, int fallback)
        {
            var node = args[name];
            if (node is null)
            {
                return fallback;
            }
            if (node.GetValueKind() != JsonValueKind.Number || !node.AsValue().TryGetValue<int>(out var n))
            {
                throw Invalid($"{name}: expected a whole number");
            }
            return n < min || n > max ? throw Invalid($"{name}: must be from {min} to {max}") : n;
        }

        public string Enum(string name, string[] values, string fallback)
        {
            var s = String(name);
            return s is null ? fallback
                : values.Contains(s) ? s
                : throw Invalid($"{name}: one of {string.Join(", ", values)}");
        }

        private ToolArgumentException Invalid(string what) => new($"Invalid arguments for {tool}: {what}");
    }
}
