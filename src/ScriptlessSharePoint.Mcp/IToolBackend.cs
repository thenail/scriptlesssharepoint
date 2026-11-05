using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Mcp;

/// <summary>A saved connection as list_connections shows it (no secrets).</summary>
public sealed record ConnectionView(string Name, string Id, string Kind, string TenantUrl, string? Account, bool Active);

/// <summary>
/// What the tools do, behind an interface so the protocol can be tested without
/// SharePoint. <c>connection</c> is a connection's name or ID; null is the active one.
/// </summary>
public interface IToolBackend
{
    IReadOnlyList<ConnectionView> ListConnections();

    /// <summary>The connection a call will use, for messages; throws if there's none.</summary>
    ConnectionView Resolve(string? connection);

    Task<Identity> WhoAmIAsync(string? connection, CancellationToken ct);

    Task<SitePage> SearchSitesAsync(string? connection, string? query, int limit, string? cursor, bool admin, CancellationToken ct);

    Task<SiteDetails> GetSiteAsync(string? connection, string url, CancellationToken ct);

    Task<IReadOnlyList<SiteAppCatalog>> ListSiteAppCatalogsAsync(string? connection, CancellationToken ct);

    Task<AppCatalogChange> EnableSiteAppCatalogAsync(string? connection, string url, CancellationToken ct);

    Task<AppCatalogChange> DisableSiteAppCatalogAsync(string? connection, string url, CancellationToken ct);

    /// <summary>Starts the Scriptless SharePoint app, showing a site when one is given.</summary>
    Task OpenAppAsync(string? siteUrl);
}
