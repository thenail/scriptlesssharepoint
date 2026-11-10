using System.Net;

using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;

using ScriptlessSharePoint.Auth;

namespace ScriptlessSharePoint.SharePoint;

/// <summary>An operation failed for a reason the user can act on; the message says what to do.</summary>
public sealed class SharePointOperationException(string message, Exception? inner = null) : Exception(message, inner);

public static class SharePointErrors
{
    /// <summary>A message for the user: what went wrong and what to do about it.</summary>
    public static string Describe(Exception e)
    {
        switch (e)
        {
            case AggregateException { InnerExceptions.Count: 1 } agg:
                return Describe(agg.InnerException!);
            case SignInRequiredException or SharePointOperationException or ArgumentException:
                return e.Message;
            case ServerUnauthorizedAccessException:
                return "Access denied. The connection doesn't have permission for this here: a site's settings need its owner, and tenant settings, app catalogs and the list of all sites need a SharePoint administrator (or an app with Sites.FullControl.All).";
            case ServerException server:
                return $"SharePoint: {server.Message}";
            case MsalUiRequiredException:
                return "Sign in again: the saved sign-in has expired or needs consent.";
            case MsalServiceException msal:
                return Msal(msal);
            case MsalClientException client when client.ErrorCode == "authentication_canceled":
                return "Sign-in was cancelled.";
            case MsalException msal:
                return $"Sign-in failed: {msal.Message}";
            case WebException { Response: HttpWebResponse response }:
                return Http((int)response.StatusCode, response.Headers["x-ms-diagnostics"]);
            case HttpRequestException { StatusCode: { } status }:
                return Http((int)status, null);
            case WebException or HttpRequestException:
                return $"Couldn't reach SharePoint: {e.Message}";
            case InvalidOperationException or NotSupportedException:
                return e.Message;
            default:
                return $"Unexpected error: {e.Message}";
        }
    }

    private static string Msal(MsalServiceException e)
    {
        var code = e.Message.Split(':', 2)[0];
        return code switch
        {
            "AADSTS65001" => "The app registration hasn't been granted the SharePoint permission it asks for. An administrator has to consent to it (Entra admin centre › App registrations › API permissions › Grant admin consent).",
            "AADSTS700016" => "There's no app with that app (client) ID in the tenant. Check the ID, and that the registration supports accounts in this organisation (multi-tenant for the built-in app).",
            "AADSTS50011" => "The app registration's redirect URIs don't include this app's. Add ms-appx-web://microsoft.aad.brokerplugin/<client id> (Windows sign-in) or http://localhost (browser sign-in) under Mobile and desktop applications.",
            "AADSTS7000215" or "AADSTS700027" => "The certificate isn't one of the app registration's credentials: upload its .cer under Certificates & secrets.",
            "AADSTS90002" or "AADSTS90023" => "The tenant ID or domain isn't right.",
            _ => $"Sign-in failed: {e.Message}",
        };
    }

    private static string Http(int status, string? diagnostics) => status switch
    {
        401 when diagnostics?.Contains("Unsupported app only token", StringComparison.OrdinalIgnoreCase) == true =>
            "SharePoint doesn't accept this app-only token: app-only access to SharePoint needs a certificate, not a client secret.",
        401 => "SharePoint rejected the token (401). Check that it's for this tenant's SharePoint (https://<tenant>.sharepoint.com), that it hasn't expired, and that the app has SharePoint permissions.",
        403 => "Access denied (403). The account or app doesn't have permission for this here: a site's settings need its owner, and tenant settings and app catalogs need a SharePoint administrator.",
        404 => "SharePoint couldn't find that site (404).",
        429 or 503 => "SharePoint is throttling requests. Wait a minute and try again.",
        _ => $"SharePoint answered HTTP {status}.",
    };

    /// <summary>
    /// How long to wait before trying again, for throttling (429) and short
    /// outages (503); null for anything else. SharePoint's Retry-After when it
    /// gives one (up to 5 minutes), else 5, 10, 20, 40, then 60 seconds.
    /// </summary>
    internal static TimeSpan? RetryAfter(Exception e, int attempt)
    {
        int? status = e switch
        {
            WebException { Response: HttpWebResponse r } => (int)r.StatusCode,
            HttpRequestException { StatusCode: { } s } => (int)s,
            _ => null,
        };
        if (status is not (429 or 503))
        {
            return null;
        }
        TimeSpan? asked = e switch
        {
            ThrottledException { RetryAfter: { } after } => after,
            WebException { Response: { } response } when int.TryParse(response.Headers["Retry-After"], out var seconds) => TimeSpan.FromSeconds(seconds),
            _ => null,
        };
        var wait = asked ?? TimeSpan.FromSeconds(Math.Min(60, 5 * Math.Pow(2, Math.Min(attempt, 4))));
        return TimeSpan.FromSeconds(Math.Clamp(wait.TotalSeconds, 1, 300));
    }
}
