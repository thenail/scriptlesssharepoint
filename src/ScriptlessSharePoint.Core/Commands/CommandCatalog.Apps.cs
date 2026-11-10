using System.Text.Json.Nodes;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>
/// Apps (SPFx solutions and add-ins) in the tenant's app catalog and site
/// collections' own: adding, deploying and removing them, and installing them
/// in sites (the PnP *-PnPApp cmdlets, through the app management REST API).
/// </summary>
public static partial class CommandCatalog
{
    private static readonly IReadOnlyList<Choice> AppCatalogs =
    [
        new("tenant", "The tenant's app catalog"),
        new("site", "The site collection's own app catalog"),
    ];

    private static CommandParameter AppName() =>
        new("app", "App", ParameterKind.Text) { Placeholder = "Contoso news ticker", Hint = "Its title, as the app catalog shows it, or its ID." };

    private static CommandParameter AppFrom() =>
        new("from", "From", ParameterKind.Choice) { Choices = AppCatalogs, Default = "tenant" };

    private static IEnumerable<Command> AppCommands()
    {
        // ---- the tenant's app catalog ----

        yield return new Command
        {
            Id = "tenant.apps",
            Title = "List apps in the app catalog",
            Group = "Apps and features",
            Description = "The apps in the tenant's app catalog: version, deployed or not, and ID.",
            PowerShell = "Get-PnPApp",
            Targets = TenantOnly,
            NeedsAdmin = true,
            ReadOnly = true,
            Glyph = "",
            Run = async (c, _) =>
            {
                var catalog = await TenantCatalogAsync(c).ConfigureAwait(false);
                var apps = await c.Client.AppsAsync(catalog, site: false, c.Cancellation).ConfigureAwait(false);
                return apps.Count == 0
                    ? CommandResult.Done("The app catalog has no apps.")
                    : CommandResult.Done($"{Plural(apps.Count, "app", "apps")}, {apps.Count(a => a.Deployed):N0} deployed.", string.Join(Environment.NewLine, apps
                        .OrderBy(a => a.Title, StringComparer.CurrentCultureIgnoreCase)
                        .Select(a => $"{a.Title}  ·  {a.CatalogVersion ?? "?"}  ·  {(a.Deployed ? "deployed" : "not deployed")}  ·  {a.Id}")));
            },
        };

        yield return new Command
        {
            Id = "tenant.app-add",
            Title = "Add an app to the app catalog",
            Group = "Apps and features",
            Description = "Uploads an app package (.sppkg) to the tenant's app catalog and, unless told not to, deploys it.",
            PowerShell = "Add-PnPApp -Publish",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Glyph = "",
            Parameters =
            [
                new("file", "App package", ParameterKind.LocalFile) { Hint = "An .sppkg (SharePoint Framework) or .app file." },
                Toggle("overwrite", "Replace the app if it's there already", true),
                Toggle("deploy", "Deploy it", true),
                Toggle("everywhere", "Make it available to every site", false, "Without installing it in each one: only for apps built for it.") with { ShowIf = new ShowIf("deploy", "true") },
            ],
            Run = async (c, _) =>
            {
                var path = c.Args.Text("file");
                var name = Path.GetFileName(path);
                if (Path.GetExtension(name).ToLowerInvariant() is not (".sppkg" or ".app"))
                {
                    return CommandResult.Failed("App packages are .sppkg or .app files.");
                }
                var catalog = await TenantCatalogAsync(c).ConfigureAwait(false);
                var bytes = await System.IO.File.ReadAllBytesAsync(path, c.Cancellation).ConfigureAwait(false);
                var id = await c.Client.AddAppAsync(catalog, site: false, name, bytes, c.Args.Toggle("overwrite"), c.Cancellation).ConfigureAwait(false);
                if (!c.Args.Toggle("deploy"))
                {
                    return CommandResult.Done($"Added {name} ({id}).");
                }
                await c.Client.AppActionAsync(catalog, site: false, id, "Deploy", DeployBody(c), c.Cancellation).ConfigureAwait(false);
                return CommandResult.Done($"Added and deployed {name} ({id}).");
            },
        };

        yield return new Command
        {
            Id = "tenant.app-deploy",
            Title = "Deploy an app",
            Group = "Apps and features",
            Description = "Deploys an app in the tenant's app catalog, so sites can install it (or, for apps built for it, every site has it).",
            PowerShell = "Publish-PnPApp",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Glyph = "",
            Parameters = [AppName(), Toggle("everywhere", "Make it available to every site", false, "Without installing it in each one: only for apps built for it.")],
            Run = (c, _) => TenantAppAsync(c, "Deploy", DeployBody(c), "Deployed"),
        };

        yield return new Command
        {
            Id = "tenant.app-retract",
            Title = "Retract an app",
            Group = "Apps and features",
            Description = "Takes an app's deployment back: sites can't install it any more, and where it's installed it stops working. It stays in the catalog.",
            PowerShell = "Unpublish-PnPApp",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Glyph = "",
            Parameters = [AppName()],
            Run = (c, _) => TenantAppAsync(c, "Retract", null, "Retracted"),
        };

        yield return new Command
        {
            Id = "tenant.app-remove",
            Title = "Remove an app from the app catalog",
            Group = "Apps and features",
            Description = "Takes an app out of the tenant's app catalog. Where it's installed, it stops working.",
            PowerShell = "Remove-PnPApp",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Destructive = true,
            Glyph = "",
            Parameters = [AppName()],
            Run = (c, _) => TenantAppAsync(c, "Remove", null, "Removed"),
        };

        // ---- in sites ----

        yield return new Command
        {
            Id = "web.apps",
            Title = "List installed apps",
            Group = "Apps and features",
            Description = "The apps installed in each site, from the tenant's app catalog and the site collection's own, and whether an update is waiting.",
            PowerShell = "Get-PnPApp",
            Targets = Webs,
            ReadOnly = true,
            Glyph = "",
            Run = async (c, n) =>
            {
                var lines = new List<string>();
                var available = 0;
                foreach (var site in new[] { false, true })
                {
                    IReadOnlyList<CatalogApp> apps;
                    try
                    {
                        apps = await c.Client.AppsAsync(n.WebUrl, site, c.Cancellation).ConfigureAwait(false);
                    }
                    catch (Exception e) when (site && e is SharePointOperationException or HttpRequestException)
                    {
                        continue; // no app catalog of its own
                    }
                    available += apps.Count(a => !a.Installed && a.Deployed);
                    lines.AddRange(apps.Where(a => a.Installed)
                        .OrderBy(a => a.Title, StringComparer.CurrentCultureIgnoreCase)
                        .Select(a => $"{a.Title}  ·  {a.InstalledVersion}{(a.CanUpgrade ? $"  ·  update to {a.CatalogVersion} waiting" : "")}  ·  {(site ? "site collection's catalog" : "tenant's catalog")}"));
                }
                return CommandResult.Done($"{Plural(lines.Count, "app", "apps")} installed, {available:N0} more to install.", string.Join(Environment.NewLine, lines));
            },
        };

        yield return SiteApp("web.app-install", "Install an app", "Installs a deployed app in each site.", "Install-PnPApp", "", "Install", false,
            a => a.Installed ? $"{a.Title} is installed already." : !a.Deployed ? $"{a.Title} isn't deployed in its catalog." : null,
            "Installing: SharePoint finishes in a minute or so.");
        yield return SiteApp("web.app-upgrade", "Update an app", "Updates an installed app in each site to the version in the catalog.", "Update-PnPApp", "", "Upgrade", false,
            a => !a.Installed ? $"{a.Title} isn't installed." : !a.CanUpgrade ? $"{a.Title} is up to date ({a.InstalledVersion})." : null,
            "Updating: SharePoint finishes in a minute or so.");
        yield return SiteApp("web.app-uninstall", "Uninstall an app", "Uninstalls an app from each site. What it added (lists, pages) may go with it.", "Uninstall-PnPApp", "", "Uninstall", true,
            a => !a.Installed ? $"{a.Title} isn't installed." : null,
            "Uninstalled.");
    }

    private static Command SiteApp(string id, string title, string description, string powerShell, string glyph, string action, bool destructive,
        Func<CatalogApp, string?> skip, string done) => new()
    {
        Id = id,
        Title = title,
        Group = "Apps and features",
        Description = description,
        PowerShell = powerShell,
        Targets = Webs,
        Destructive = destructive,
        Glyph = glyph,
        Parameters = [AppName(), AppFrom()],
        Run = async (c, n) =>
        {
            var site = c.Args.Text("from") == "site";
            var apps = await c.Client.AppsAsync(n.WebUrl, site, c.Cancellation).ConfigureAwait(false);
            if (FindApp(apps, c.Args.Text("app")) is not { } app)
            {
                return CommandResult.Failed(NoSuchApp(apps, c.Args.Text("app")));
            }
            if (skip(app) is { } reason)
            {
                return CommandResult.Skipped(reason);
            }
            await c.Client.AppActionAsync(n.WebUrl, site, app.Id, action, null, c.Cancellation).ConfigureAwait(false);
            return CommandResult.Done(done);
        },
    };

    private static async Task<string> TenantCatalogAsync(CommandContext c) =>
        await c.Client.TenantAppCatalogUrlAsync(c.Cancellation).ConfigureAwait(false)
        ?? throw new SharePointOperationException("The tenant has no app catalog yet: create one in the SharePoint admin centre (More features › Apps).");

    private static async Task<CommandResult> TenantAppAsync(CommandContext c, string action, JsonObject? body, string done)
    {
        var catalog = await TenantCatalogAsync(c).ConfigureAwait(false);
        var apps = await c.Client.AppsAsync(catalog, site: false, c.Cancellation).ConfigureAwait(false);
        if (FindApp(apps, c.Args.Text("app")) is not { } app)
        {
            return CommandResult.Failed(NoSuchApp(apps, c.Args.Text("app")));
        }
        await c.Client.AppActionAsync(catalog, site: false, app.Id, action, body, c.Cancellation).ConfigureAwait(false);
        return CommandResult.Done($"{done} {app.Title}.");
    }

    private static JsonObject DeployBody(CommandContext c) => new() { ["skipFeatureDeployment"] = c.Args.Toggle("everywhere") };

    /// <summary>An app by its ID, or else its title (any case); null when none or more than one has it.</summary>
    internal static CatalogApp? FindApp(IReadOnlyList<CatalogApp> apps, string typed)
    {
        var text = typed.Trim();
        if (Guid.TryParse(text, out var id))
        {
            return apps.FirstOrDefault(a => a.Id == id);
        }
        var named = apps.Where(a => string.Equals(a.Title, text, StringComparison.CurrentCultureIgnoreCase)).ToList();
        return named.Count == 1 ? named[0] : null;
    }

    private static string NoSuchApp(IReadOnlyList<CatalogApp> apps, string typed)
    {
        var text = typed.Trim();
        return apps.Count(a => string.Equals(a.Title, text, StringComparison.CurrentCultureIgnoreCase)) > 1
            ? $"More than one app is called {text}: give its ID (List apps shows them)."
            : $"The catalog has no app called {text}.";
    }
}
