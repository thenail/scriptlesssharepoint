using System.Diagnostics;
using System.IO;
using System.Text.Json;

using ScriptlessSharePoint.Auth;
using ScriptlessSharePoint.Connections;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Mcp;

/// <summary>
/// The tools' backend: the app's saved connections (re-read on every call, so
/// changes in the app apply at once), signing in silently only, and CSOM.
/// </summary>
public sealed class SharePointBackend(ConnectionStore store) : IToolBackend
{
    private readonly Dictionary<string, (string Fingerprint, SharePointClient Client)> clients = [];

    public IReadOnlyList<ConnectionView> ListConnections()
    {
        var data = store.Load();
        return data.Connections.Select(c => View(c, c.Id == data.Active?.Id)).ToList();
    }

    public ConnectionView Resolve(string? connection)
    {
        var data = store.Load();
        var profile = Find(data, connection);
        return View(profile, profile.Id == data.Active?.Id);
    }

    public Task<Identity> WhoAmIAsync(string? connection, CancellationToken ct) => Client(connection).WhoAmIAsync(ct);

    public Task<SitePage> SearchSitesAsync(string? connection, string? query, int limit, string? cursor, bool admin, CancellationToken ct) =>
        admin
            ? Client(connection).ListSitesAsAdminAsync(query, limit, cursor, ct)
            : Client(connection).SearchSitesAsync(query, limit, cursor, ct);

    public Task<SiteDetails> GetSiteAsync(string? connection, string url, CancellationToken ct) => Client(connection).GetSiteAsync(url, ct);

    public Task<IReadOnlyList<SiteAppCatalog>> ListSiteAppCatalogsAsync(string? connection, CancellationToken ct) =>
        Client(connection).ListSiteAppCatalogsAsync(ct);

    public Task<AppCatalogChange> EnableSiteAppCatalogAsync(string? connection, string url, CancellationToken ct) =>
        Client(connection).EnableSiteAppCatalogAsync(url, ct);

    public Task<AppCatalogChange> DisableSiteAppCatalogAsync(string? connection, string url, CancellationToken ct) =>
        Client(connection).DisableSiteAppCatalogAsync(url, ct);

    /// <summary>Starts the app: the one next to this exe, else its execution alias.</summary>
    public Task OpenAppAsync(string? siteUrl)
    {
        var sibling = Path.Combine(AppContext.BaseDirectory, AppInfo.AppExe);
        var start = new ProcessStartInfo(File.Exists(sibling) ? sibling : AppInfo.AppAlias)
        {
            UseShellExecute = true,
        };
        if (!string.IsNullOrWhiteSpace(siteUrl))
        {
            start.ArgumentList.Add("--site");
            start.ArgumentList.Add(siteUrl);
        }
        try
        {
            Process.Start(start)?.Dispose();
        }
        catch (Exception e) when (e is System.ComponentModel.Win32Exception or InvalidOperationException)
        {
            throw new SharePointOperationException(
                $"Couldn't start {AppInfo.Name}. It may not be installed: get it from the Microsoft Store.", e);
        }
        return Task.CompletedTask;
    }

    /// <summary>The connection's client, made again when the connection changed in the app.</summary>
    private SharePointClient Client(string? connection)
    {
        var profile = Find(store.Load(), connection);
        if (profile.Validate() is { } problem)
        {
            throw new SharePointOperationException($"The connection \"{profile.Name}\" isn't complete: {problem} Fix it in {AppInfo.Name}.");
        }
        var fingerprint = JsonSerializer.Serialize(profile);
        if (clients.TryGetValue(profile.Id, out var cached) && cached.Fingerprint == fingerprint)
        {
            return cached.Client;
        }
        // silent sign-in only: no window, and accounts are saved by the app
        var client = new SharePointClient(profile, TokenProviders.Create(profile), interactive: false);
        clients[profile.Id] = (fingerprint, client);
        return client;
    }

    private static ConnectionProfile Find(ConnectionData data, string? connection)
    {
        if (data.Connections.Count == 0)
        {
            throw new SharePointOperationException(
                $"There are no connections yet. Ask the user to open {AppInfo.Name} (open_app) and add one.");
        }
        return data.Find(connection) ?? throw new SharePointOperationException(
            $"There's no connection \"{connection}\". The connections are: {string.Join(", ", data.Connections.Select(c => c.Name))}.");
    }

    private static ConnectionView View(ConnectionProfile c, bool active) =>
        new(c.Name, c.Id, c.KindLabel, c.TenantUrl, c.Kind == ConnectionKind.User ? c.AccountName : null, active);
}
