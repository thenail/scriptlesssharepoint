namespace ScriptlessSharePoint.Auth;

/// <summary>Gets access tokens for SharePoint resources (https://contoso.sharepoint.com, …-admin…).</summary>
public interface ITokenProvider
{
    /// <summary>
    /// A token for <paramref name="resource"/>. With <paramref name="interactive"/>
    /// false (the MCP server) it never shows UI and throws
    /// <see cref="SignInRequiredException"/> when the user has to sign in.
    /// </summary>
    Task<string> GetTokenAsync(string resource, bool interactive, CancellationToken cancellationToken = default);

    /// <summary>Forgets the signed-in account, where there is one.</summary>
    Task SignOutAsync();
}

/// <summary>The connection can't get a token without the user: sign in in the app.</summary>
public sealed class SignInRequiredException(string message, Exception? inner = null) : Exception(message, inner);
