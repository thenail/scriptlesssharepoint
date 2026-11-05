using System.Text;
using System.Text.Json;

using ScriptlessSharePoint.Connections;

namespace ScriptlessSharePoint.Auth;

/// <summary>A pasted access token, sent as it is; it can't be refreshed.</summary>
public sealed class AccessTokenProvider(ConnectionProfile profile) : ITokenProvider
{
    public Task<string> GetTokenAsync(string resource, bool interactive, CancellationToken cancellationToken = default)
    {
        var token = Secrets.Unprotect(profile.ProtectedAccessToken)
            ?? throw new SignInRequiredException(
                $"The access token of \"{profile.Name}\" can't be read on this PC or by this user. Paste it again in {AppInfo.Name}.");
        var info = TokenInfo.TryParse(token);
        if (info?.Expires is { } expires && expires <= DateTimeOffset.UtcNow.AddSeconds(30))
        {
            throw new SignInRequiredException(
                $"The access token of \"{profile.Name}\" expired at {expires.ToLocalTime():g}. Paste a new one in {AppInfo.Name}.");
        }
        return Task.FromResult(token);
    }

    public Task SignOutAsync() => Task.CompletedTask;
}

/// <summary>What an access token says about itself (its claims are read, not verified).</summary>
public sealed record TokenInfo(string? Audience, DateTimeOffset? Expires, string? Subject, string? TenantId, string? Permissions)
{
    /// <summary>The SharePoint Online resource's app ID, an audience some tokens carry instead of the URL.</summary>
    public const string SharePointAppId = "00000003-0000-0ff1-ce00-000000000000";

    public bool IsForSharePoint =>
        Audience is not null &&
        (Audience.Equals(SharePointAppId, StringComparison.OrdinalIgnoreCase) ||
         Audience.Contains(".sharepoint.", StringComparison.OrdinalIgnoreCase));

    public static TokenInfo? TryParse(string? token)
    {
        var parts = token?.Trim().Split('.');
        if (parts is not { Length: 3 })
        {
            return null;
        }
        try
        {
            var payload = parts[1].Replace('-', '+').Replace('_', '/');
            payload = payload.PadRight(payload.Length + ((4 - (payload.Length % 4)) % 4), '=');
            using var doc = JsonDocument.Parse(Encoding.UTF8.GetString(Convert.FromBase64String(payload)));
            var root = doc.RootElement;
            string? Str(string name) => root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;
            DateTimeOffset? expires = root.TryGetProperty("exp", out var exp) && exp.TryGetInt64(out var seconds)
                ? DateTimeOffset.FromUnixTimeSeconds(seconds)
                : null;
            var roles = root.TryGetProperty("roles", out var r) && r.ValueKind == JsonValueKind.Array
                ? string.Join(' ', r.EnumerateArray().Select(x => x.GetString()))
                : null;
            return new TokenInfo(
                Str("aud"),
                expires,
                Str("upn") ?? Str("unique_name") ?? Str("app_displayname") ?? Str("appid"),
                Str("tid"),
                Str("scp") ?? roles);
        }
        catch (Exception e) when (e is FormatException or JsonException or ArgumentException)
        {
            return null;
        }
    }
}
