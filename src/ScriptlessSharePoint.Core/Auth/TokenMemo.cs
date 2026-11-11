namespace ScriptlessSharePoint.Auth;

/// <summary>
/// Access tokens already got, by resource, until shortly before they expire:
/// every SharePoint request asks for one, and asking Windows' broker each time
/// costs a round trip (seconds when it's busy).
/// </summary>
public sealed class TokenMemo
{
    /// <summary>How long before its expiry a token stops being handed out: long enough for a request (and its retries) to finish with it.</summary>
    public static readonly TimeSpan Margin = TimeSpan.FromMinutes(5);

    private readonly Dictionary<string, (string Token, DateTimeOffset ExpiresOn)> tokens = new(StringComparer.OrdinalIgnoreCase);
    private readonly object gate = new();

    /// <summary>The token kept for <paramref name="resource"/>, or null when there's none still good at <paramref name="now"/>.</summary>
    public string? Get(string resource, DateTimeOffset now)
    {
        lock (gate)
        {
            if (!tokens.TryGetValue(Key(resource), out var kept))
            {
                return null;
            }
            if (kept.ExpiresOn - Margin > now)
            {
                return kept.Token;
            }
            tokens.Remove(Key(resource));
            return null;
        }
    }

    public void Set(string resource, string token, DateTimeOffset expiresOn)
    {
        lock (gate)
        {
            tokens[Key(resource)] = (token, expiresOn);
        }
    }

    /// <summary>Forgets every token: signed out, or signed in as someone else.</summary>
    public void Clear()
    {
        lock (gate)
        {
            tokens.Clear();
        }
    }

    private static string Key(string resource) => resource.TrimEnd('/');
}
