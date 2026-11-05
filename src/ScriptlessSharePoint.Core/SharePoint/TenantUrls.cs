using System.Text;
using System.Text.RegularExpressions;

namespace ScriptlessSharePoint.SharePoint;

/// <summary>SharePoint Online URLs: the tenant root, its admin centre and site URLs.</summary>
public static partial class TenantUrls
{
    /// <summary>
    /// "contoso", "contoso.sharepoint.com" or any URL in the tenant
    /// → "https://contoso.sharepoint.com". Admin and OneDrive hosts map to the root.
    /// </summary>
    public static bool TryNormalizeTenantUrl(string? input, out string tenantUrl)
    {
        tenantUrl = "";
        var s = input?.Trim() ?? "";
        if (s.Length == 0)
        {
            return false;
        }
        if (TenantName().IsMatch(s))
        {
            tenantUrl = $"https://{s.ToLowerInvariant()}.sharepoint.com";
            return true;
        }
        if (!s.Contains("://", StringComparison.Ordinal))
        {
            s = "https://" + s;
        }
        if (!Uri.TryCreate(s, UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeHttps || !uri.Host.Contains('.'))
        {
            return false;
        }
        var host = uri.Host.ToLowerInvariant();
        var dot = host.IndexOf('.');
        var label = host[..dot];
        foreach (var suffix in new[] { "-admin", "-my" })
        {
            if (label.EndsWith(suffix, StringComparison.Ordinal))
            {
                label = label[..^suffix.Length];
            }
        }
        tenantUrl = $"https://{label}{host[dot..]}";
        return true;
    }

    public static string NormalizeTenantUrl(string input) =>
        TryNormalizeTenantUrl(input, out var url)
            ? url
            : throw new ArgumentException($"\"{input}\" isn't a SharePoint tenant (e.g. contoso or https://contoso.sharepoint.com).");

    /// <summary>https://contoso.sharepoint.com → https://contoso-admin.sharepoint.com.</summary>
    public static string AdminUrl(string tenantUrl)
    {
        if (!TryNormalizeTenantUrl(tenantUrl, out var root))
        {
            return "";
        }
        var uri = new Uri(root);
        var dot = uri.Host.IndexOf('.');
        return $"https://{uri.Host[..dot]}-admin{uri.Host[dot..]}";
    }

    /// <summary>The resource (audience) a token for a URL must be for: its scheme and host.</summary>
    public static string Resource(string url) => new Uri(url).GetLeftPart(UriPartial.Authority);

    private static readonly string[] PageMarkers =
    [
        "/_layouts/", "/_api/", "/_vti_bin/", "/SitePages/", "/Lists/", "/Forms/",
        "/Shared Documents", "/Shared%20Documents", "/SiteAssets/", "/Pages/",
    ];

    /// <summary>
    /// A site (web) URL from what a user or agent typed: an absolute URL in the
    /// tenant, or a server-relative path ("/sites/hr"). Page, list and library
    /// parts are cut off, and so are the query and fragment.
    /// </summary>
    public static string NormalizeSiteUrl(string input, string tenantUrl)
    {
        var s = input?.Trim() ?? "";
        if (s.Length == 0)
        {
            throw new ArgumentException("Give a site URL, e.g. https://contoso.sharepoint.com/sites/hr.");
        }
        var root = NormalizeTenantUrl(tenantUrl);
        if (s.StartsWith('/'))
        {
            s = root + s;
        }
        else if (!s.Contains("://", StringComparison.Ordinal))
        {
            s = "https://" + s;
        }
        if (!Uri.TryCreate(s, UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeHttps)
        {
            throw new ArgumentException($"\"{input}\" isn't an https:// SharePoint URL.");
        }
        var rootHost = new Uri(root).Host;
        if (!string.Equals(uri.Host, rootHost, StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException(
                $"{uri.Host} isn't this connection's tenant ({rootHost}). Use a site URL on {rootHost}, or a connection for that tenant.");
        }
        var path = Uri.UnescapeDataString(uri.AbsolutePath);
        foreach (var marker in PageMarkers)
        {
            var at = path.IndexOf(Uri.UnescapeDataString(marker), StringComparison.OrdinalIgnoreCase);
            if (at >= 0)
            {
                path = path[..at];
            }
        }
        if (path.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase))
        {
            path = path[..path.LastIndexOf('/')];
        }
        path = string.Join('/', path.TrimEnd('/').Split('/').Select(Uri.EscapeDataString));
        return $"https://{uri.Host.ToLowerInvariant()}{path}";
    }

    /// <summary>
    /// The KQL query for site collections whose title or URL has words starting
    /// with <paramref name="text"/>'s words; OneDrive sites are left out.
    /// </summary>
    public static string SiteSearchQuery(string? text)
    {
        var sb = new StringBuilder();
        foreach (var word in SearchWords(text))
        {
            // lower case: AND, OR, NOT and NEAR are operators only in capitals
            sb.Append(word.ToLowerInvariant()).Append("* ");
        }
        sb.Append("contentclass:STS_Site -WebTemplate:SPSPERS");
        return sb.ToString();
    }

    /// <summary>The words of a search (letters and digits; at most eight), as typed.</summary>
    public static IReadOnlyList<string> SearchWords(string? text) =>
        [.. SearchWord().Matches(text ?? "").Select(m => m.Value).Take(8)];

    [GeneratedRegex("^[a-zA-Z0-9][a-zA-Z0-9-]{0,62}$")]
    private static partial Regex TenantName();

    /// <summary>Letters and digits only: KQL operators and quotes in the text can't change the query.</summary>
    [GeneratedRegex(@"[\p{L}\p{N}]+")]
    private static partial Regex SearchWord();
}
