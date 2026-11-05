namespace ScriptlessSharePoint.SharePoint;

/// <summary>A site collection in search or admin results.</summary>
public sealed record SiteSummary(
    string Url,
    string Title,
    string? Template,
    DateTimeOffset? LastModified,
    Guid? SiteId);

/// <summary>
/// A page of sites. <see cref="Next"/> is the cursor for the next page (null at
/// the end). <see cref="Total"/> counts every match when the source knows it
/// (search does; admin searches don't, and the admin listing's is an estimate:
/// <see cref="TotalIsEstimate"/>). Source: "search" or "admin".
/// </summary>
public sealed record SitePage(
    IReadOnlyList<SiteSummary> Sites,
    string? Next,
    int? Total,
    bool TotalIsEstimate,
    string Source);

/// <summary>Whether a site collection has its own app catalog.</summary>
public enum AppCatalogState
{
    Enabled,
    NotEnabled,
    Unknown,
}

/// <summary>
/// A site collection and its app catalog. <see cref="AppCatalogSource"/> is
/// "tenant" when the state comes from the tenant's list of site collection app
/// catalogs (needs an administrator), "site" when it comes from the site having
/// an Apps for SharePoint library. <see cref="RequestedUrl"/> is the URL asked
/// about, when it was a subsite or page of <see cref="Url"/>.
/// </summary>
public sealed record SiteDetails(
    string Url,
    string Title,
    string WebTemplate,
    Guid SiteId,
    AppCatalogState AppCatalog,
    string AppCatalogSource,
    string? AppCatalogUrl,
    int? AppCount,
    string? RequestedUrl);

/// <summary>A site collection app catalog, as the tenant lists them.</summary>
public sealed record SiteAppCatalog(string Url, Guid SiteId, string? Error);

/// <summary>What enabling or removing a site collection app catalog did.</summary>
public sealed record AppCatalogChange(string SiteUrl, bool Changed, string Message);

/// <summary>Who a connection acts as, and the tenant root site's title.</summary>
public sealed record Identity(string DisplayName, string? Email, string? LoginName, bool IsApp, string TenantTitle);
