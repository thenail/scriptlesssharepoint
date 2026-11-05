using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;

namespace ScriptlessSharePoint.Auth;

/// <summary>
/// The MSAL token cache on disk (DPAPI-encrypted) in <see cref="AppInfo.DataDir"/>,
/// shared by the app and the MCP server: who signs in in the app is signed in for AI apps.
/// With Windows' broker (WAM) Windows keeps the tokens and this keeps the accounts.
/// </summary>
public static class MsalCache
{
    private static readonly SemaphoreSlim Gate = new(1, 1);
    private static MsalCacheHelper? helper;

    public static async Task RegisterAsync(ITokenCache cache)
    {
        await Gate.WaitAsync().ConfigureAwait(false);
        try
        {
            if (helper is null)
            {
                Directory.CreateDirectory(AppInfo.DataDir);
                var props = new StorageCreationPropertiesBuilder("msal.cache", AppInfo.DataDir).Build();
                helper = await MsalCacheHelper.CreateAsync(props).ConfigureAwait(false);
            }
            helper.RegisterCache(cache);
        }
        finally
        {
            Gate.Release();
        }
    }
}
