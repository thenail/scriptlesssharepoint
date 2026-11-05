using ScriptlessSharePoint.Connections;

namespace ScriptlessSharePoint.Auth;

public static class TokenProviders
{
    /// <summary>The token provider for a connection.</summary>
    /// <param name="parentWindow">The window sign-in dialogs belong to; null where there's no UI (the MCP server).</param>
    /// <param name="accountChanged">Called when a sign-in changes the connection's account, to save it.</param>
    public static ITokenProvider Create(
        ConnectionProfile profile,
        Func<IntPtr>? parentWindow = null,
        Action<ConnectionProfile>? accountChanged = null) =>
        profile.Kind switch
        {
            ConnectionKind.User => new DelegatedTokenProvider(profile, parentWindow, accountChanged),
            ConnectionKind.Certificate => new CertificateTokenProvider(profile),
            ConnectionKind.AccessToken => new AccessTokenProvider(profile),
            _ => throw new NotSupportedException($"Unknown connection kind {profile.Kind}."),
        };
}
