using System.Text.Json.Serialization;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Connections;

/// <summary>How a connection gets its SharePoint access tokens.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<ConnectionKind>))]
public enum ConnectionKind
{
    /// <summary>A person signs in (delegated): the Windows account or another one.</summary>
    User,

    /// <summary>App-only: an Entra app with a certificate (SharePoint rejects client secrets).</summary>
    Certificate,

    /// <summary>An access token pasted by the user; used as-is until it expires.</summary>
    AccessToken,
}

/// <summary>One saved SharePoint Online connection.</summary>
public sealed class ConnectionProfile
{
    public string Id { get; set; } = Guid.NewGuid().ToString("n");

    public string Name { get; set; } = "";

    public ConnectionKind Kind { get; set; } = ConnectionKind.User;

    /// <summary>The tenant's root, e.g. https://contoso.sharepoint.com.</summary>
    public string TenantUrl { get; set; } = "";

    /// <summary>The admin centre, when it isn't the usual https://contoso-admin.sharepoint.com.</summary>
    public string? AdminUrl { get; set; }

    /// <summary>The Entra app (client) ID; null uses the built-in one (User connections only).</summary>
    public string? ClientId { get; set; }

    /// <summary>The Entra tenant (ID or domain). Required for Certificate; User defaults to "organizations".</summary>
    public string? TenantId { get; set; }

    /// <summary>User: sign in silently as the account signed in to Windows.</summary>
    public bool UseWindowsAccount { get; set; } = true;

    /// <summary>User: sign in in the browser instead of Windows' account broker (WAM),
    /// for app registrations without the broker's redirect URI.</summary>
    public bool UseBrowser { get; set; }

    /// <summary>User: the MSAL account last signed in (HomeAccountId.Identifier).</summary>
    public string? AccountId { get; set; }

    /// <summary>User: that account's user name, shown and used as the login hint.</summary>
    public string? AccountName { get; set; }

    /// <summary>Certificate: thumbprint of a certificate with a private key in CurrentUser\My or LocalMachine\My.</summary>
    public string? CertificateThumbprint { get; set; }

    /// <summary>AccessToken: the token, encrypted for this Windows user (DPAPI).</summary>
    public string? ProtectedAccessToken { get; set; }

    [JsonIgnore]
    public string EffectiveClientId => string.IsNullOrWhiteSpace(ClientId) ? AppInfo.BuiltInClientId ?? "" : ClientId.Trim();

    [JsonIgnore]
    public string EffectiveAdminUrl => string.IsNullOrWhiteSpace(AdminUrl) ? TenantUrls.AdminUrl(TenantUrl) : AdminUrl.Trim().TrimEnd('/');

    [JsonIgnore]
    public string KindLabel => Kind switch
    {
        ConnectionKind.User when UseWindowsAccount => "Windows account",
        ConnectionKind.User => "Signed-in user",
        ConnectionKind.Certificate => "App-only (certificate)",
        ConnectionKind.AccessToken => "Access token",
        _ => Kind.ToString(),
    };

    /// <summary>A shallow copy, for editing without touching the saved one.</summary>
    public ConnectionProfile Clone() => (ConnectionProfile)MemberwiseClone();

    /// <summary>What's missing before this connection can be used, or null.</summary>
    public string? Validate()
    {
        if (string.IsNullOrWhiteSpace(Name))
        {
            return "Give the connection a name.";
        }
        if (!TenantUrls.TryNormalizeTenantUrl(TenantUrl, out _))
        {
            return "Enter the SharePoint tenant, e.g. contoso or https://contoso.sharepoint.com.";
        }
        switch (Kind)
        {
            case ConnectionKind.User:
                if (!Guid.TryParse(EffectiveClientId, out _))
                {
                    return "Enter the app (client) ID of an Entra app registration.";
                }
                if (UseWindowsAccount && UseBrowser)
                {
                    return "The Windows account signs in through Windows, not the browser: turn one of them off.";
                }
                break;
            case ConnectionKind.Certificate:
                if (!Guid.TryParse(ClientId, out _))
                {
                    return "Enter the app (client) ID of the Entra app registration.";
                }
                if (string.IsNullOrWhiteSpace(TenantId))
                {
                    return "App-only sign-in needs the tenant ID or domain (e.g. contoso.onmicrosoft.com).";
                }
                if (string.IsNullOrWhiteSpace(CertificateThumbprint))
                {
                    return "Choose a certificate: import a .pfx, create one, or enter a thumbprint.";
                }
                break;
            case ConnectionKind.AccessToken:
                if (string.IsNullOrWhiteSpace(ProtectedAccessToken))
                {
                    return "Paste an access token for SharePoint.";
                }
                break;
        }
        return null;
    }

    public override string ToString() => Name;
}
