using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Broker;

using ScriptlessSharePoint.Connections;
using ScriptlessSharePoint.Diagnostics;

namespace ScriptlessSharePoint.Auth;

/// <summary>
/// A person signs in (delegated permissions): the account signed in to Windows,
/// silently through Windows' account broker (WAM), or any account through WAM's
/// account picker or, for app registrations without WAM's redirect URI, the browser.
/// </summary>
public sealed class DelegatedTokenProvider : ITokenProvider
{
    private readonly ConnectionProfile profile;
    private readonly IPublicClientApplication app;
    private readonly Action<ConnectionProfile>? accountChanged;
    private readonly SemaphoreSlim gate = new(1, 1);
    /// <summary>The UI thread's context when made by the app: sign-in dialogs open there.</summary>
    private readonly SynchronizationContext? ui = SynchronizationContext.Current;
    private readonly TokenMemo memo = new();
    private bool cacheReady;
    /// <summary>The saved account, once looked up: looking it up asks WAM for every account (up to seconds), so once is enough.</summary>
    private IAccount? savedAccount;

    /// <param name="parentWindow">The window WAM's dialogs belong to; null where there's no UI.</param>
    /// <param name="accountChanged">Called after an interactive sign-in, to save the account.</param>
    public DelegatedTokenProvider(ConnectionProfile profile, Func<IntPtr>? parentWindow, Action<ConnectionProfile>? accountChanged)
    {
        this.profile = profile;
        this.accountChanged = accountChanged;
        var tenant = string.IsNullOrWhiteSpace(profile.TenantId) ? "organizations" : profile.TenantId.Trim();
        var builder = PublicClientApplicationBuilder.Create(profile.EffectiveClientId)
            .WithAuthority(AzureCloudInstance.AzurePublic, tenant)
            .WithClientName(AppInfo.Name)
            .WithClientVersion(AppInfo.Version)
            // warnings only: Info writes some 270 lines per sign-in, each a file append on MSAL's path
            .WithLogging(Log.Msal, LogLevel.Warning, enablePiiLogging: false, enableDefaultPlatformLogging: false);
        builder = profile.UseBrowser
            ? builder.WithRedirectUri("http://localhost")
            : builder.WithBroker(new BrokerOptions(BrokerOptions.OperatingSystems.Windows)
            {
                Title = AppInfo.Name,
                ListOperatingSystemAccounts = true,
            });
        if (parentWindow is not null)
        {
            builder = builder.WithParentActivityOrWindow(parentWindow);
        }
        app = builder.Build();
    }

    public async Task<string> GetTokenAsync(string resource, bool interactive, CancellationToken cancellationToken = default)
    {
        string[] scopes = [$"{resource.TrimEnd('/')}/.default"];
        if (memo.Get(resource, DateTimeOffset.UtcNow) is { } kept)
        {
            return kept;
        }
        await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            // another request may have got it while this one waited
            if (memo.Get(resource, DateTimeOffset.UtcNow) is { } got)
            {
                return got;
            }
            if (!cacheReady)
            {
                await MsalCache.RegisterAsync(app.UserTokenCache).ConfigureAwait(false);
                cacheReady = true;
            }

            // each account that may sign in silently, in turn: one Windows won't
            // use (e.g. signed out of the app) doesn't mean the next can't
            var watch = System.Diagnostics.Stopwatch.StartNew();
            var candidates = await SilentAccountsAsync().ConfigureAwait(false);
            var lookup = watch.ElapsedMilliseconds;
            var how = profile.UseBrowser ? "browser" : "WAM";
            MsalUiRequiredException? lastFailure = null;
            foreach (var (account, who) in candidates)
            {
                watch.Restart();
                try
                {
                    var silent = await app.AcquireTokenSilent(scopes, account).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                    Log.Write("auth", $"silent token for {resource} as {who} ({how}): ok, from {silent.AuthenticationResultMetadata.TokenSource} " +
                                      $"in {watch.ElapsedMilliseconds} ms (accounts {lookup} ms)");
                    return Keep(resource, silent);
                }
                catch (MsalUiRequiredException e)
                {
                    Log.Write("auth", $"silent token for {resource} as {who} ({how}): needs sign-in ({e.ErrorCode}) in {watch.ElapsedMilliseconds} ms {FirstLine(e.Message)}");
                    lastFailure = e;
                }
            }
            if (!interactive)
            {
                if (candidates.Count == 0)
                {
                    Log.Write("auth", $"silent token for {resource}: no account to sign in silently as");
                }
                throw new SignInRequiredException(NeedsSignIn(), lastFailure);
            }

            var request = app.AcquireTokenInteractive(scopes);
            var first = candidates.Count > 0 ? candidates[0].Account : null;
            if (profile.UseWindowsAccount && first is not null)
            {
                request = request.WithAccount(first);
            }
            else
            {
                request = request.WithPrompt(Prompt.SelectAccount);
                if (!string.IsNullOrWhiteSpace(profile.AccountName))
                {
                    request = request.WithLoginHint(profile.AccountName);
                }
            }
            var mode = profile.UseWindowsAccount && first is not null ? "with the Windows account" : "with the account picker";
            Log.Write("auth", $"interactive sign-in for {resource} {mode} ({how}, tenant {(string.IsNullOrWhiteSpace(profile.TenantId) ? "organizations" : "given")}): started");
            AuthenticationResult result;
            try
            {
                result = await OnUiThread(() => request.ExecuteAsync(cancellationToken)).ConfigureAwait(false);
            }
            catch (MsalException e)
            {
                Log.Write("auth", $"interactive sign-in for {resource}: failed ({e.ErrorCode}) {FirstLine(e.Message)}");
                throw;
            }
            Log.Write("auth", $"interactive sign-in for {resource}: ok, tenant {result.TenantId}, scopes {string.Join(' ', result.Scopes)}");
            return Keep(resource, result);
        }
        finally
        {
            gate.Release();
        }
    }

    public async Task SignOutAsync()
    {
        await gate.WaitAsync().ConfigureAwait(false);
        try
        {
            if (!cacheReady)
            {
                await MsalCache.RegisterAsync(app.UserTokenCache).ConfigureAwait(false);
                cacheReady = true;
            }
            memo.Clear();
            savedAccount = null;
            // only this connection's account: never every account the app knows
            if (profile.AccountId is not null &&
                await app.GetAccountAsync(profile.AccountId).ConfigureAwait(false) is { } account)
            {
                await app.RemoveAsync(account).ConfigureAwait(false);
                Log.Write("auth", "signed out the connection's account");
            }
            profile.AccountId = null;
            profile.AccountName = null;
            accountChanged?.Invoke(profile);
        }
        finally
        {
            gate.Release();
        }
    }

    /// <summary>The accounts to try silently, in order: the saved one, then (Windows account connections) Windows' own.</summary>
    private async Task<List<(IAccount Account, string Who)>> SilentAccountsAsync()
    {
        var list = new List<(IAccount, string)>();
        if (!string.IsNullOrEmpty(profile.AccountId))
        {
            if (savedAccount?.HomeAccountId?.Identifier != profile.AccountId)
            {
                savedAccount = await app.GetAccountAsync(profile.AccountId).ConfigureAwait(false);
            }
            if (savedAccount is not null)
            {
                list.Add((savedAccount, "the saved account"));
            }
        }
        if (profile.UseWindowsAccount && !profile.UseBrowser)
        {
            list.Add((PublicClientApplication.OperatingSystemAccount, "the Windows account"));
        }
        return list;
    }

    private static string FirstLine(string s) => s.Split('\n', 2)[0].Trim();

    private Task<T> OnUiThread<T>(Func<Task<T>> run)
    {
        if (ui is null || SynchronizationContext.Current == ui)
        {
            return run();
        }
        var done = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        ui.Post(async _ =>
        {
            try
            {
                done.SetResult(await run());
            }
            catch (Exception e)
            {
                done.SetException(e);
            }
        }, null);
        return done.Task;
    }

    /// <summary>Notes who a token is for and keeps it for the next requests.</summary>
    private string Keep(string resource, AuthenticationResult result)
    {
        Remember(result.Account);
        memo.Set(resource, result.AccessToken, result.ExpiresOn);
        return result.AccessToken;
    }

    private void Remember(IAccount? account)
    {
        var id = account?.HomeAccountId?.Identifier;
        if (id is null)
        {
            return;
        }
        if (id != profile.AccountId)
        {
            // someone else now: the tokens kept were theirs
            memo.Clear();
        }
        savedAccount = account;
        if (id == profile.AccountId && account!.Username == profile.AccountName)
        {
            return;
        }
        profile.AccountId = id;
        profile.AccountName = account!.Username;
        accountChanged?.Invoke(profile);
    }

    private string NeedsSignIn() =>
        $"The connection \"{profile.Name}\" isn't signed in. Open {AppInfo.Name} and sign in, then try again.";
}
