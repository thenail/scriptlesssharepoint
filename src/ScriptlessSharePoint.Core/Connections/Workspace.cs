using System.Text.Json;

using ScriptlessSharePoint.Auth;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Connections;

/// <summary>A failure the app shows as it is: what's wrong and what to do.</summary>
public sealed class ConnectionException(string message) : Exception(message);

/// <summary>
/// The app's connections and the session of the one being used: saving, testing
/// and deleting connections, signing in and out. One per window, made on the UI
/// thread, so interactive sign-in (WAM) opens over that window.
/// </summary>
public sealed class Workspace
{
    private readonly Func<IntPtr> window;
    /// <summary>Accounts signed in with by <see cref="TestAsync"/>, applied when that connection is saved.</summary>
    private readonly Dictionary<string, (string? Id, string? Name)> testedAccounts = [];
    private Session? session;

    /// <param name="window">The window sign-in dialogs belong to; read from MSAL's threads, so a cached handle.</param>
    public Workspace(ConnectionStore store, Func<IntPtr> window)
    {
        Store = store;
        this.window = window;
    }

    public ConnectionStore Store { get; }

    /// <summary>A sign-in changed a connection's account (saved already). May be raised off the UI thread.</summary>
    public event Action<ConnectionProfile>? AccountChanged;

    public ConnectionData Load() => Store.Load();

    // ---- connections ----

    /// <summary>Saves an edited connection and makes it the active one.</summary>
    public ConnectionProfile Save(ConnectionDraft draft, string? token)
    {
        ConnectionProfile? saved = null;
        Store.Update(data =>
        {
            var existing = data.Connections.FirstOrDefault(c => c.Id == draft.Id);
            var profile = draft.ToProfile(existing, token);
            if (testedAccounts.Remove(profile.Id, out var account) && profile.Kind == ConnectionKind.User)
            {
                (profile.AccountId, profile.AccountName) = account;
            }
            if (profile.Validate() is { } problem)
            {
                throw new ConnectionException(problem);
            }
            if (existing is null)
            {
                data.Connections.Add(profile);
            }
            else
            {
                data.Connections[data.Connections.IndexOf(existing)] = profile;
            }
            data.ActiveId = profile.Id;
            saved = profile;
        });
        session = null; // the next call starts a session with the saved connection
        return saved!;
    }

    /// <summary>
    /// Removes a connection. Its account stays signed in: other connections (and
    /// AI apps) may use it, and Windows would refuse it silently until signed in again.
    /// </summary>
    public ConnectionData Delete(string id)
    {
        var profile = Store.Load().Find(id) ?? throw new ConnectionException("That connection doesn't exist any more.");
        var data = Store.Update(d =>
        {
            d.Connections.RemoveAll(c => c.Id == profile.Id);
            if (d.ActiveId == profile.Id)
            {
                d.ActiveId = d.Connections.FirstOrDefault()?.Id;
            }
        });
        if (session?.Profile.Id == profile.Id)
        {
            session = null;
        }
        return data;
    }

    public void SetActive(string id) => Store.Update(d =>
    {
        if (d.Find(id) is null)
        {
            throw new ConnectionException("That connection doesn't exist any more.");
        }
        d.ActiveId = id;
    });

    /// <summary>Signs in with a connection as edited (not yet saved) and says who it is.</summary>
    public async Task<TestResult> TestAsync(ConnectionDraft draft, string? token)
    {
        var saved = draft.Id is null ? null : Store.Load().Connections.FirstOrDefault(c => c.Id == draft.Id);
        var profile = draft.ToProfile(saved, token);
        if (profile.Validate() is { } problem)
        {
            throw new ConnectionException(problem);
        }
        var tokens = TokenProviders.Create(profile, window);
        var identity = await new SharePointClient(profile, tokens, interactive: true).WhoAmIAsync();
        await PrepareAdminTokenAsync(profile, tokens);
        testedAccounts[profile.Id] = (profile.AccountId, profile.AccountName);
        return new TestResult(profile.Id, identity, profile.AccountName);
    }

    // ---- the session ----

    /// <summary>The active connection's session: made on first use, again when the connection changed.</summary>
    public Session Current()
    {
        var data = Store.Load();
        var profile = data.Active ?? throw new ConnectionException("Add a connection first.");
        if (profile.Validate() is { } problem)
        {
            throw new ConnectionException($"The connection \"{profile.Name}\" isn't complete: {problem}");
        }
        var fingerprint = JsonSerializer.Serialize(profile);
        if (session is null || session.Profile.Id != profile.Id || session.Fingerprint != fingerprint)
        {
            session = new Session(profile, fingerprint, window, OnAccountChanged);
        }
        return session;
    }

    /// <summary>Who the active connection is signed in as, found silently; null when it needs signing in.</summary>
    public async Task<Identity?> StatusAsync()
    {
        var s = Current();
        if (s.Identity is null)
        {
            try
            {
                s.Identity = await s.Silent.WhoAmIAsync();
            }
            catch (Exception)
            {
                // not signed in (or offline): the app offers Sign in
            }
        }
        return s.Identity;
    }

    public async Task<Identity> SignInAsync()
    {
        var s = Current();
        s.Identity = await s.Client.WhoAmIAsync();
        await PrepareAdminTokenAsync(s.Profile, s.Tokens);
        return s.Identity;
    }

    public async Task SignOutAsync()
    {
        var s = Current();
        await s.Tokens.SignOutAsync();
        s.Identity = null;
    }

    /// <summary>
    /// The admin centre is its own resource (https://contoso-admin.sharepoint.com):
    /// get its token while the user is signing in, so a prompt for it, if Windows
    /// needs one, comes now rather than at the first admin operation. Its failure
    /// isn't the sign-in's: the admin operations report it when they're used.
    /// </summary>
    private static async Task PrepareAdminTokenAsync(ConnectionProfile profile, ITokenProvider tokens)
    {
        if (profile.Kind != ConnectionKind.User)
        {
            return;
        }
        try
        {
            await tokens.GetTokenAsync(TenantUrls.Resource(profile.EffectiveAdminUrl), interactive: true);
        }
        catch (Exception e)
        {
            Diagnostics.Log.Write("auth", $"admin centre token after sign-in: {e.GetType().Name}: {SharePointErrors.Describe(e)}");
        }
    }

    private void OnAccountChanged(ConnectionProfile profile)
    {
        Store.Update(d =>
        {
            var saved = d.Connections.FirstOrDefault(c => c.Id == profile.Id);
            if (saved is not null)
            {
                saved.AccountId = profile.AccountId;
                saved.AccountName = profile.AccountName;
            }
        });
        if (session?.Profile.Id == profile.Id)
        {
            // the account is saved now; the session stays the one signed in
            session.Fingerprint = JsonSerializer.Serialize(Store.Load().Find(profile.Id));
        }
        AccountChanged?.Invoke(profile);
    }
}

/// <summary>What testing a connection found: its ID (new connections get one), who it is, and the account.</summary>
public sealed record TestResult(string Id, Identity Identity, string? AccountName);

/// <summary>A connection in use: its tokens, who it's signed in as, and its SharePoint clients.</summary>
public sealed class Session(ConnectionProfile profile, string fingerprint, Func<IntPtr> window, Action<ConnectionProfile> accountChanged)
{
    private SharePointClient? client;
    private SharePointClient? silent;

    public ConnectionProfile Profile { get; } = profile;
    internal string Fingerprint { get; set; } = fingerprint;
    public ITokenProvider Tokens { get; } = TokenProviders.Create(profile, window, accountChanged);
    public Identity? Identity { get; set; }

    /// <summary>Signs in interactively when it has to.</summary>
    public SharePointClient Client => client ??= new SharePointClient(Profile, Tokens, interactive: true);

    /// <summary>Never shows sign-in UI.</summary>
    public SharePointClient Silent => silent ??= new SharePointClient(Profile, Tokens, interactive: false);
}
