using System.Text;

using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>Site collections, as the admin centre manages them (Set-SPOSite and friends).</summary>
public static partial class CommandCatalog
{
    private static IEnumerable<Command> SiteCollection()
    {
        // ---- settings ----

        yield return new Command
        {
            Id = "site.storage",
            Title = "Storage limit",
            Group = "Settings",
            Description = "How much each site collection may store, and when its owners get a warning. Only where the tenant sets storage by hand.",
            PowerShell = "Set-SPOSite -StorageQuota -StorageQuotaWarningLevel",
            Targets = Sites,
            NeedsAdmin = true,
            Glyph = "\uEDA2",
            Parameters =
            [
                new("max", "Limit (MB)", ParameterKind.Number) { Min = 1, Max = 25 * 1024 * 1024, Placeholder = "102400", Hint = "1 GB is 1024 MB; 100 GB is 102400." },
                new("warn", "Warn at (MB)", ParameterKind.Number) { Required = false, Min = 1, Max = 25 * 1024 * 1024, Hint = "Empty: 90 % of the limit." },
            ],
            Run = (c, n) => c.UpdateSiteAsync(n, p =>
            {
                var max = c.Args.Number("max");
                p.StorageMaximumLevel = max;
                p.StorageWarningLevel = c.Args["warn"] is null ? max * 9 / 10 : Math.Min(c.Args.Number("warn"), max);
            }, $"Storage limit set to {c.Args.Number("max"):N0} MB."),
        };

        yield return new Command
        {
            Id = "site.lock",
            Title = "Lock or unlock",
            Group = "Settings",
            Description = "Read-only stops changes; no access shuts everyone out (they see the tenant's no-access page).",
            PowerShell = "Set-SPOSite -LockState",
            Targets = Sites,
            NeedsAdmin = true,
            Destructive = true,
            Glyph = "\uE72E",
            Parameters =
            [
                new("state", "State", ParameterKind.Choice)
                {
                    Default = "Unlock",
                    Choices = [new("Unlock", "Unlocked"), new("ReadOnly", "Read-only"), new("NoAccess", "No access")],
                },
            ],
            Run = (c, n) => c.UpdateSiteAsync(n, p => p.LockState = c.Args.Text("state"), c.Args.Text("state") switch
            {
                "Unlock" => "Unlocked.",
                "ReadOnly" => "Made read-only.",
                _ => "Locked: no access.",
            }),
        };

        yield return new Command
        {
            Id = "site.scripts",
            Title = "Custom scripts",
            Group = "Settings",
            Description = "Whether owners may add custom script (and change the property bag, and use some classic features). Microsoft turns it off again after 24 hours.",
            PowerShell = "Set-SPOSite -DenyAddAndCustomizePages",
            Targets = Sites,
            NeedsAdmin = true,
            Glyph = "\uE943",
            Parameters = [Toggle("allow", "Allow custom scripts", false)],
            Run = (c, n) => c.UpdateSiteAsync(n,
                p => p.DenyAddAndCustomizePages = c.Args.Toggle("allow") ? DenyAddAndCustomizePagesStatus.Disabled : DenyAddAndCustomizePagesStatus.Enabled,
                c.Args.Toggle("allow") ? "Custom scripts allowed." : "Custom scripts blocked."),
        };

        yield return new Command
        {
            Id = "site.access-policy",
            Title = "Access from unmanaged devices",
            Group = "Settings",
            Description = "Full access, web-only access (no downloading, printing or syncing), or none, for devices Conditional Access doesn't manage.",
            PowerShell = "Set-SPOSite -ConditionalAccessPolicy",
            Targets = Sites,
            NeedsAdmin = true,
            Glyph = "\uE8EA",
            Parameters =
            [
                new("policy", "Unmanaged devices get", ParameterKind.Choice)
                {
                    Default = nameof(SPOConditionalAccessPolicyType.AllowFullAccess),
                    Choices =
                    [
                        new(nameof(SPOConditionalAccessPolicyType.AllowFullAccess), "Full access"),
                        new(nameof(SPOConditionalAccessPolicyType.AllowLimitedAccess), "Limited, web-only access"),
                        new(nameof(SPOConditionalAccessPolicyType.BlockAccess), "No access"),
                    ],
                },
            ],
            Run = (c, n) => c.UpdateSiteAsync(n, p => p.ConditionalAccessPolicy = c.Args.Enum<SPOConditionalAccessPolicyType>("policy"), "Changed access from unmanaged devices."),
        };

        yield return new Command
        {
            Id = "site.page-comments",
            Title = "Comments on pages",
            Group = "Settings",
            Description = "Whether people can comment on the site's pages.",
            PowerShell = "Set-SPOSite -CommentsOnSitePagesDisabled",
            Targets = Sites,
            NeedsAdmin = true,
            Glyph = "\uE90A",
            Parameters = [Toggle("allow", "Allow comments", true)],
            Run = (c, n) => c.UpdateSiteAsync(n, p => p.CommentsOnSitePagesDisabled = !c.Args.Toggle("allow"),
                c.Args.Toggle("allow") ? "Comments allowed." : "Comments turned off."),
        };

        // ---- sharing ----

        yield return new Command
        {
            Id = "site.sharing",
            Title = "Site sharing",
            Group = "Sharing",
            Description = "Who the site's content can be shared with. It can't be looser than the tenant's sharing.",
            PowerShell = "Set-SPOSite -SharingCapability",
            Targets = Sites,
            NeedsAdmin = true,
            Glyph = "\uE72D",
            Parameters = [new("level", "Share with", ParameterKind.Choice) { Choices = SharingLevels, Default = nameof(SharingCapabilities.ExistingExternalUserSharingOnly) }],
            Run = (c, n) => c.UpdateSiteAsync(n, p => p.SharingCapability = c.Args.Enum<SharingCapabilities>("level"), "Changed the site's sharing."),
        };

        yield return new Command
        {
            Id = "site.default-link",
            Title = "Default sharing link",
            Group = "Sharing",
            Description = "The kind of link, and what it allows, that Share offers first in these sites.",
            PowerShell = "Set-SPOSite -DefaultSharingLinkType -DefaultLinkPermission",
            Targets = Sites,
            NeedsAdmin = true,
            Glyph = "\uE71B",
            Parameters =
            [
                new("type", "Link", ParameterKind.Choice) { Choices = LinkTypes, Default = nameof(SharingLinkType.None) },
                new("permission", "Allows", ParameterKind.Choice) { Choices = LinkPermissions, Default = nameof(SharingPermissionType.None) },
            ],
            Run = (c, n) => c.UpdateSiteAsync(n, p =>
            {
                p.DefaultSharingLinkType = c.Args.Enum<SharingLinkType>("type");
                p.DefaultLinkPermission = c.Args.Enum<SharingPermissionType>("permission");
            }, "Changed the default link."),
        };

        yield return new Command
        {
            Id = "site.company-links",
            Title = "\"People in your organisation\" links",
            Group = "Sharing",
            Description = "Whether people can share with links that work for everyone in the organisation.",
            PowerShell = "Set-SPOSite -DisableCompanyWideSharingLinks",
            Targets = Sites,
            NeedsAdmin = true,
            Glyph = "\uE716",
            Parameters = [Toggle("allow", "Allow these links", true)],
            Run = (c, n) => c.UpdateSiteAsync(n,
                p => p.DisableCompanyWideSharingLinks = c.Args.Toggle("allow") ? CompanyWideSharingLinksPolicy.NotDisabled : CompanyWideSharingLinksPolicy.Disabled,
                c.Args.Toggle("allow") ? "Organisation links allowed." : "Organisation links turned off."),
        };

        // ---- administrators ----

        yield return new Command
        {
            Id = "site.owner",
            Title = "Change the primary administrator",
            Group = "Administrators",
            Description = "Makes someone the site collection's owner (primary administrator).",
            PowerShell = "Set-SPOSite -Owner",
            Targets = Sites,
            NeedsAdmin = true,
            Glyph = "\uE77B",
            Parameters = [Person("New owner", "Their email address or user name.")],
            Run = (c, n) => c.UpdateSiteAsync(n, p => p.Owner = c.Args.Text("who"), $"{c.Args.Text("who")} is the owner now."),
        };

        yield return new Command
        {
            Id = "site.admin-add",
            Title = "Add a site collection administrator",
            Group = "Administrators",
            Description = "Gives someone full control of everything in the site collection.",
            PowerShell = "Set-SPOUser -IsSiteCollectionAdmin $true",
            Targets = Sites,
            Glyph = "\uE8FA",
            Parameters = [Person()],
            Run = (c, n) => SetSiteAdminAsync(c, n, c.Args.Text("who"), true),
        };

        yield return new Command
        {
            Id = "site.admin-remove",
            Title = "Remove a site collection administrator",
            Group = "Administrators",
            Description = "Takes someone off the site collection's administrators. The primary administrator can't be removed this way.",
            PowerShell = "Set-SPOUser -IsSiteCollectionAdmin $false",
            Targets = Sites,
            Glyph = "\uE8AB",
            Parameters = [Person()],
            Run = (c, n) => SetSiteAdminAsync(c, n, c.Args.Text("who"), false),
        };

        yield return new Command
        {
            Id = "site.admins",
            Title = "List site collection administrators",
            Group = "Administrators",
            Description = "Who has full control of each site collection.",
            PowerShell = "Get-SPOUser | Where IsSiteAdmin",
            Targets = Sites,
            ReadOnly = true,
            Glyph = "\uE716",
            Run = async (c, n) =>
            {
                using var ctx = await c.SiteAsync(n).ConfigureAwait(false);
                var admins = ctx.LoadQuery(ctx.Web.SiteUsers.Where(u => u.IsSiteAdmin).Include(u => u.Title, u => u.Email, u => u.LoginName));
                await ctx.RunAsync().ConfigureAwait(false);
                var list = admins.Select(u => $"{u.Title}  ·  {(string.IsNullOrEmpty(u.Email) ? u.LoginName : u.Email)}").ToList();
                return CommandResult.Done($"{list.Count} administrators.", string.Join(Environment.NewLine, list));
            },
        };

        // ---- hub sites ----

        yield return new Command
        {
            Id = "site.hub-register",
            Title = "Make it a hub site",
            Group = "Hub sites",
            Description = "Registers each site collection as a hub site that other sites can join.",
            PowerShell = "Register-SPOHubSite",
            Targets = Sites,
            NeedsAdmin = true,
            Glyph = "\uE8D4",
            Run = async (c, n) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                new Tenant(ctx).RegisterHubSite(n.Url);
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("It's a hub site now.");
            },
        };

        yield return new Command
        {
            Id = "site.hub-unregister",
            Title = "Stop being a hub site",
            Group = "Hub sites",
            Description = "Unregisters the hub site. The sites that had joined it leave it.",
            PowerShell = "Unregister-SPOHubSite",
            Targets = Sites,
            NeedsAdmin = true,
            Destructive = true,
            Glyph = "\uE8D4",
            Run = async (c, n) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                new Tenant(ctx).UnregisterHubSite(n.Url);
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("It isn't a hub site any more.");
            },
        };

        yield return new Command
        {
            Id = "site.hub-join",
            Title = "Join a hub site",
            Group = "Hub sites",
            Description = "Associates each site collection with a hub site: its navigation, theme and search.",
            PowerShell = "Add-SPOHubSiteAssociation",
            Targets = Sites,
            NeedsAdmin = true,
            Glyph = "\uE71B",
            Parameters = [new("hub", "Hub site address", ParameterKind.Url) { Placeholder = "https://contoso.sharepoint.com/sites/intranet" }],
            Run = async (c, n) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                new Tenant(ctx).ConnectSiteToHubSite(n.Url, c.Args.Text("hub").TrimEnd('/'));
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Joined the hub site.");
            },
        };

        yield return new Command
        {
            Id = "site.hub-leave",
            Title = "Leave its hub site",
            Group = "Hub sites",
            Description = "Takes each site collection out of the hub site it has joined.",
            PowerShell = "Remove-SPOHubSiteAssociation",
            Targets = Sites,
            NeedsAdmin = true,
            Glyph = "\uE8BB",
            Run = async (c, n) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                new Tenant(ctx).DisconnectSiteFromHubSite(n.Url);
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Left the hub site.");
            },
        };

        // ---- apps and features ----

        yield return new Command
        {
            Id = "site.app-catalog-add",
            Title = "Add a site collection app catalog",
            Group = "Apps and features",
            Description = "Gives each site collection its own app catalog, for apps only it uses.",
            PowerShell = "Add-SPOSiteCollectionAppCatalog",
            Targets = Sites,
            NeedsAdmin = true,
            Glyph = "\uE71D",
            Run = async (c, n) =>
            {
                var r = await c.Client.EnableSiteAppCatalogAsync(n.Url, c.Cancellation).ConfigureAwait(false);
                return r.Changed ? CommandResult.Done(r.Message) : CommandResult.Skipped(r.Message);
            },
        };

        yield return new Command
        {
            Id = "site.app-catalog-remove",
            Title = "Remove the site collection app catalog",
            Group = "Apps and features",
            Description = "Stops the site collection's app catalog being one. Its library and files stay.",
            PowerShell = "Remove-SPOSiteCollectionAppCatalog",
            Targets = Sites,
            NeedsAdmin = true,
            Glyph = "\uE71D",
            Run = async (c, n) =>
            {
                var r = await c.Client.DisableSiteAppCatalogAsync(n.Url, c.Cancellation).ConfigureAwait(false);
                return r.Changed ? CommandResult.Done(r.Message) : CommandResult.Skipped(r.Message);
            },
        };

        foreach (var command in FeatureCommands("site", Sites, "Site", (ctx, _) => ctx.Site.Features))
        {
            yield return command;
        }

        // ---- recycle bin ----

        yield return new Command
        {
            Id = "site.recycle-restore",
            Title = "Restore everything in the recycle bin",
            Group = "Recycle bin",
            Description = "Puts back everything deleted in the site collection (both stages of its recycle bin).",
            PowerShell = "Restore-PnPRecycleBinItem -All",
            Targets = Sites,
            Glyph = "\uE7A7",
            Run = async (c, n) =>
            {
                using var ctx = await c.SiteAsync(n).ConfigureAwait(false);
                ctx.Site.RecycleBin.RestoreAll();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Restored everything in the recycle bin.");
            },
        };

        yield return new Command
        {
            Id = "site.recycle-empty",
            Title = "Empty the recycle bin",
            Group = "Recycle bin",
            Description = "Deletes everything in the site collection's recycle bin (both stages) for good.",
            PowerShell = "Clear-PnPRecycleBinItem -All",
            Targets = Sites,
            Destructive = true,
            Glyph = "\uE74D",
            Run = async (c, n) =>
            {
                using var ctx = await c.SiteAsync(n).ConfigureAwait(false);
                ctx.Site.RecycleBin.DeleteAll();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Emptied the recycle bin.");
            },
        };

        // ---- delete ----

        yield return new Command
        {
            Id = "site.delete",
            Title = "Delete site collection",
            Group = "Delete",
            Description = "Moves each site collection to the tenant's recycle bin, where it stays for 93 days.",
            PowerShell = "Remove-SPOSite",
            Targets = Sites,
            NeedsAdmin = true,
            Destructive = true,
            After = After.Remove,
            Glyph = "\uE74D",
            Run = async (c, n) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                var op = new Tenant(ctx).RemoveSite(n.Url);
                ctx.Load(op, o => o.IsComplete, o => o.PollingInterval);
                await ctx.RunAsync().ConfigureAwait(false);
                await Csom.WaitAsync(ctx, op, c.Cancellation).ConfigureAwait(false);
                return CommandResult.Done("Deleted (in the tenant's recycle bin).");
            },
        };
    }

    /// <summary>
    /// Adds or removes a site collection administrator: through the admin centre
    /// (Set-SPOUser), or, for a site owner who isn't a SharePoint administrator,
    /// in the site itself.
    /// </summary>
    private static async Task<CommandResult> SetSiteAdminAsync(CommandContext c, SpNode n, string who, bool admin)
    {
        try
        {
            using var ctx = await c.AdminAsync().ConfigureAwait(false);
            new Tenant(ctx).SetSiteAdmin(n.Url, who, admin);
            await ctx.RunAsync().ConfigureAwait(false);
        }
        catch (Exception e) when (SharePointClient.IsAccessDenied(e))
        {
            using var ctx = await c.SiteAsync(n).ConfigureAwait(false);
            var user = Csom.EnsureUser(ctx, who);
            user.IsSiteAdmin = admin;
            user.Update();
            await ctx.RunAsync().ConfigureAwait(false);
        }
        return CommandResult.Done(admin ? $"{who} is an administrator." : $"{who} isn't an administrator any more.");
    }

    /// <summary>Turning a feature on and off, by its ID, at site collection or web scope.</summary>
    private static IEnumerable<Command> FeatureCommands(string prefix, NodeKind[] targets, string scope, Func<ClientContext, SpNode, FeatureCollection> features)
    {
        var id = new CommandParameter("feature", "Feature ID", ParameterKind.Text)
        {
            Placeholder = "00000000-0000-0000-0000-000000000000",
            Hint = $"The feature's GUID ({scope.ToLowerInvariant()}-scoped).",
        };
        Guid Parse(CommandContext c) =>
            Guid.TryParse(c.Args.Text("feature"), out var g) ? g : throw new ArgumentException("The feature ID isn't a GUID.");

        yield return new Command
        {
            Id = $"{prefix}.feature-on",
            Title = scope == "Site" ? "Turn on a site collection feature" : "Turn on a site feature",
            Group = "Apps and features",
            Description = $"Activates a {scope.ToLowerInvariant()}-scoped feature.",
            PowerShell = $"Enable-PnPFeature -Scope {scope}",
            Targets = targets,
            Glyph = "\uEA86",
            Parameters = [id],
            Run = async (c, n) =>
            {
                using var ctx = await (scope == "Site" ? c.SiteAsync(n) : c.WebAsync(n)).ConfigureAwait(false);
                features(ctx, n).Add(Parse(c), false, FeatureDefinitionScope.None);
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Turned the feature on.");
            },
        };
        yield return new Command
        {
            Id = $"{prefix}.feature-off",
            Title = scope == "Site" ? "Turn off a site collection feature" : "Turn off a site feature",
            Group = "Apps and features",
            Description = $"Deactivates a {scope.ToLowerInvariant()}-scoped feature.",
            PowerShell = $"Disable-PnPFeature -Scope {scope}",
            Targets = targets,
            Glyph = "\uEA86",
            Parameters = [id],
            Run = async (c, n) =>
            {
                using var ctx = await (scope == "Site" ? c.SiteAsync(n) : c.WebAsync(n)).ConfigureAwait(false);
                features(ctx, n).Remove(Parse(c), true);
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Turned the feature off.");
            },
        };
    }

    private static string Lines(IEnumerable<string> lines) => string.Join(Environment.NewLine, lines);

    private static string Join(StringBuilder text) => text.ToString().TrimEnd();
}
