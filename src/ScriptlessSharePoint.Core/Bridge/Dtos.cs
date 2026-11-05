using ScriptlessSharePoint.Connections;

namespace ScriptlessSharePoint.Bridge;

/// <summary>
/// A connection as the web UI sees and edits it: no secrets. The token comes in
/// separately (<c>token</c>) and is never sent back; <see cref="HasToken"/> says one is saved.
/// </summary>
public sealed record ConnectionDto
{
    public string? Id { get; init; }
    public string Name { get; init; } = "";
    public ConnectionKind Kind { get; init; } = ConnectionKind.User;
    public string TenantUrl { get; init; } = "";
    public string? AdminUrl { get; init; }
    public string? ClientId { get; init; }
    public string? TenantId { get; init; }
    public bool UseWindowsAccount { get; init; } = true;
    public bool UseBrowser { get; init; }
    public string? AccountName { get; init; }
    public string? CertificateThumbprint { get; init; }
    public bool HasToken { get; init; }
    public string? KindLabel { get; init; }

    public static ConnectionDto From(ConnectionProfile p) => new()
    {
        Id = p.Id,
        Name = p.Name,
        Kind = p.Kind,
        TenantUrl = p.TenantUrl,
        AdminUrl = p.AdminUrl,
        ClientId = p.ClientId,
        TenantId = p.TenantId,
        UseWindowsAccount = p.UseWindowsAccount,
        UseBrowser = p.UseBrowser,
        AccountName = p.AccountName,
        CertificateThumbprint = p.CertificateThumbprint,
        HasToken = !string.IsNullOrEmpty(p.ProtectedAccessToken),
        KindLabel = p.KindLabel,
    };

    /// <summary>
    /// The edited profile: this DTO's fields over <paramref name="saved"/> (the
    /// stored one, or null for a new connection), keeping what the UI doesn't
    /// edit (the signed-in account, the encrypted token) unless the kind changed.
    /// </summary>
    public ConnectionProfile ToProfile(ConnectionProfile? saved, string? token)
    {
        var p = saved?.Clone() ?? new ConnectionProfile();
        if (Id is { Length: > 0 } && saved is null)
        {
            p.Id = Id;
        }
        var kindChanged = saved is not null && saved.Kind != Kind;
        p.Name = Name.Trim();
        p.Kind = Kind;
        p.TenantUrl = SharePoint.TenantUrls.TryNormalizeTenantUrl(TenantUrl, out var url) ? url : TenantUrl.Trim();
        p.AdminUrl = Blank(AdminUrl)?.TrimEnd('/');
        p.ClientId = Blank(ClientId);
        p.TenantId = Blank(TenantId);
        p.UseWindowsAccount = Kind == ConnectionKind.User && UseWindowsAccount;
        p.UseBrowser = Kind == ConnectionKind.User && !UseWindowsAccount && UseBrowser;
        p.CertificateThumbprint = Kind == ConnectionKind.Certificate && Blank(CertificateThumbprint) is { } tp
            ? Auth.Certificates.Normalize(tp)
            : null;
        if (Kind != ConnectionKind.User || kindChanged)
        {
            p.AccountId = null;
            p.AccountName = null;
        }
        if (Kind != ConnectionKind.AccessToken)
        {
            p.ProtectedAccessToken = null;
        }
        else if (CleanToken(token) is { Length: > 0 } t)
        {
            p.ProtectedAccessToken = Secrets.Protect(t);
        }
        return p;
    }

    /// <summary>A pasted token without "Bearer " and white space.</summary>
    public static string CleanToken(string? text)
    {
        var t = text?.Trim() ?? "";
        if (t.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
        {
            t = t[7..];
        }
        return string.Concat(t.Where(c => !char.IsWhiteSpace(c)));
    }

    private static string? Blank(string? s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();
}
