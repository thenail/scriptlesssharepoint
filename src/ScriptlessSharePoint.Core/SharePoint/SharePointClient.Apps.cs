using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json.Nodes;

namespace ScriptlessSharePoint.SharePoint;

/// <summary>An app (SharePoint add-in or SPFx solution) in an app catalog, as a site sees it.</summary>
public sealed record CatalogApp(
    Guid Id,
    string Title,
    string? CatalogVersion,
    string? InstalledVersion,
    bool Deployed,
    bool CanUpgrade,
    bool ClientSide)
{
    public bool Installed => !string.IsNullOrEmpty(InstalledVersion);
}

/// <summary>
/// Apps in app catalogs, through SharePoint's app management (ALM) REST API,
/// which CSOM doesn't have: the tenant's catalog, or a site collection's own.
/// Requests carry the connection's bearer token, so they need no request digest.
/// </summary>
public sealed partial class SharePointClient
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(5) };

    /// <summary>The REST path of an app catalog, seen from a web.</summary>
    private static string Catalog(bool site) => site ? "_api/web/sitecollectionappcatalog" : "_api/web/tenantappcatalog";

    /// <summary>The tenant app catalog's site, or null when the tenant has none.</summary>
    public async Task<string?> TenantAppCatalogUrlAsync(CancellationToken ct = default)
    {
        var settings = await RestAsync(TenantUrl, HttpMethod.Get, "_api/SP_TenantSettings_Current", null, ct).ConfigureAwait(false);
        return settings?["CorporateCatalogUrl"]?.GetValue<string>() is { Length: > 0 } url ? url.TrimEnd('/') : null;
    }

    /// <summary>The apps in a catalog, with their state in <paramref name="webUrl"/> (installed or not, upgradable).</summary>
    public async Task<IReadOnlyList<CatalogApp>> AppsAsync(string webUrl, bool site, CancellationToken ct = default)
    {
        var json = await RestAsync(webUrl, HttpMethod.Get, $"{Catalog(site)}/AvailableApps", null, ct).ConfigureAwait(false);
        return ParseApps(json);
    }

    /// <summary>Adds an app package to the tenant's catalog (or, with <paramref name="site"/>, the site collection's); returns its ID.</summary>
    public async Task<Guid> AddAppAsync(string catalogWebUrl, bool site, string fileName, byte[] package, bool overwrite, CancellationToken ct = default)
    {
        var content = new ByteArrayContent(package);
        content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
        var path = $"{Catalog(site)}/Add(overwrite={(overwrite ? "true" : "false")},url='{Uri.EscapeDataString(fileName.Replace("'", "''"))}')";
        var file = await RestAsync(catalogWebUrl, HttpMethod.Post, path, content, ct).ConfigureAwait(false);
        return Guid.TryParse(file?["UniqueId"]?.GetValue<string>(), out var id)
            ? id
            : throw new SharePointOperationException("SharePoint took the package but didn't say what it became.");
    }

    /// <summary>
    /// An app action: Deploy (with <paramref name="body"/>), Retract, Remove,
    /// Install, Upgrade or Uninstall, in <paramref name="webUrl"/>.
    /// </summary>
    public Task AppActionAsync(string webUrl, bool site, Guid app, string action, JsonObject? body = null, CancellationToken ct = default)
    {
        HttpContent? content = null;
        if (body is not null)
        {
            content = new StringContent(body.ToJsonString(), Encoding.UTF8);
            content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=nometadata");
        }
        return RestAsync(webUrl, HttpMethod.Post, $"{Catalog(site)}/AvailableApps/GetById('{app}')/{action}", content, ct);
    }

    /// <summary>
    /// A REST call with the connection's token, answered in JSON without
    /// metadata. Throttling and access denied throw <see cref="HttpRequestException"/>
    /// (retried, or described, like CSOM's); other refusals carry SharePoint's message.
    /// </summary>
    internal async Task<JsonNode?> RestAsync(string webUrl, HttpMethod method, string path, HttpContent? content, CancellationToken ct)
    {
        var token = await tokens.GetTokenAsync(TenantUrls.Resource(webUrl), Interactive, ct).ConfigureAwait(false);
        using var request = new HttpRequestMessage(method, $"{webUrl.TrimEnd('/')}/{path.TrimStart('/')}") { Content = content };
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        request.Headers.Accept.ParseAdd("application/json;odata=nometadata");
        request.Headers.TryAddWithoutValidation("User-Agent", $"NONISV|ScriptlessSharePoint|ScriptlessSharePoint/{AppInfo.Version}");
        if (method == HttpMethod.Post && content is null)
        {
            request.Content = new ByteArrayContent([]);
        }
        using var response = await Http.SendAsync(request, ct).ConfigureAwait(false);
        var text = await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            var message = RestError(text);
            if (response.StatusCode is HttpStatusCode.TooManyRequests or HttpStatusCode.ServiceUnavailable)
            {
                // how long SharePoint asked to wait: seconds, or a time
                var after = response.Headers.RetryAfter is { } r
                    ? r.Delta ?? (r.Date is { } date ? date - DateTimeOffset.UtcNow : null)
                    : null;
                throw new ThrottledException(message, response.StatusCode, after);
            }
            if (response.StatusCode is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden)
            {
                throw new HttpRequestException(message, null, response.StatusCode);
            }
            throw new SharePointOperationException(message is null ? $"SharePoint answered HTTP {(int)response.StatusCode}." : $"SharePoint: {message}");
        }
        return string.IsNullOrWhiteSpace(text) ? null : JsonNode.Parse(text);
    }

    /// <summary>SharePoint's own words in a REST error, in either of its JSON shapes.</summary>
    internal static string? RestError(string body)
    {
        try
        {
            var json = JsonNode.Parse(body);
            var error = json?["odata.error"] ?? json?["error"];
            return error?["message"]?["value"]?.GetValue<string>() ?? error?["message"]?.GetValue<string>();
        }
        catch (Exception e) when (e is System.Text.Json.JsonException or InvalidOperationException)
        {
            return null;
        }
    }

    internal static IReadOnlyList<CatalogApp> ParseApps(JsonNode? json)
    {
        var items = json?["value"]?.AsArray() ?? json as JsonArray ?? [];
        return items.OfType<JsonObject>().Select(a => new CatalogApp(
                Guid.TryParse(Text(a, "ID") ?? Text(a, "Id"), out var id) ? id : Guid.Empty,
                Text(a, "Title") ?? "",
                Text(a, "AppCatalogVersion"),
                Text(a, "InstalledVersion"),
                Flag(a, "Deployed"),
                Flag(a, "CanUpgrade"),
                Flag(a, "IsClientSideSolution")))
            .Where(a => a.Id != Guid.Empty)
            .ToList();

        static string? Text(JsonObject o, string name) => o[name] is JsonValue v && v.TryGetValue<string>(out var s) && s.Length > 0 ? s : null;
        static bool Flag(JsonObject o, string name) => o[name] is JsonValue v && v.TryGetValue<bool>(out var b) && b;
    }
}
