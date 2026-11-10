using System.Text;

using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>The tenant's own: new and deleted site collections, and the tenant's sharing.</summary>
public static partial class CommandCatalog
{
    private static readonly IReadOnlyList<Choice> SharingLevels =
    [
        new(nameof(SharingCapabilities.ExternalUserAndGuestSharing), "Anyone (links that need no sign-in)"),
        new(nameof(SharingCapabilities.ExternalUserSharingOnly), "New and existing guests"),
        new(nameof(SharingCapabilities.ExistingExternalUserSharingOnly), "Existing guests only"),
        new(nameof(SharingCapabilities.Disabled), "Only people in your organisation"),
    ];

    private static readonly IReadOnlyList<Choice> LinkTypes =
    [
        new(nameof(SharingLinkType.None), "The tenant's default"),
        new(nameof(SharingLinkType.AnonymousAccess), "Anyone with the link"),
        new(nameof(SharingLinkType.Internal), "People in your organisation"),
        new(nameof(SharingLinkType.Direct), "Specific people"),
    ];

    private static readonly IReadOnlyList<Choice> LinkPermissions =
    [
        new(nameof(SharingPermissionType.None), "The tenant's default"),
        new(nameof(SharingPermissionType.View), "View"),
        new(nameof(SharingPermissionType.Edit), "Edit"),
    ];

    private static IEnumerable<Command> Tenant()
    {
        yield return new Command
        {
            Id = "tenant.create-site",
            Title = "New site collection",
            Group = "Create",
            Description = "Creates a communication site or a team site without a Microsoft 365 group, and waits until SharePoint has made it.",
            PowerShell = "New-SPOSite",
            Targets = TenantOnly,
            NeedsAdmin = true,
            After = After.ReloadSites,
            Glyph = "\uE710",
            Parameters =
            [
                new("title", "Title", ParameterKind.Text),
                new("url", "Address", ParameterKind.Url) { Placeholder = "https://contoso.sharepoint.com/sites/projects", Hint = "The new site's full address, under /sites/ or /teams/." },
                new("template", "Kind of site", ParameterKind.Choice)
                {
                    Default = "SITEPAGEPUBLISHING#0",
                    Choices = [new("SITEPAGEPUBLISHING#0", "Communication site"), new("STS#3", "Team site (no Microsoft 365 group)")],
                },
                Person("Owner", "Who owns the site: their email address or user name."),
                new("quota", "Storage limit (MB)", ParameterKind.Number) { Required = false, Min = 1, Max = 25 * 1024 * 1024, Hint = "Only where the tenant sets storage by hand." },
                new("lcid", "Language (LCID)", ParameterKind.Number) { Default = "1033", Min = 1025, Max = 1106, Hint = "1033 is English. A site's language can't be changed later." },
            ],
            Run = async (c, _) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                var tenant = new Tenant(ctx);
                var props = new SiteCreationProperties
                {
                    Title = c.Args.Text("title"),
                    Url = c.Args.Text("url").TrimEnd('/'),
                    Template = c.Args.Text("template"),
                    Owner = c.Args.Text("who"),
                    Lcid = (uint)c.Args.Number("lcid"),
                };
                if (c.Args["quota"] is not null)
                {
                    props.StorageMaximumLevel = c.Args.Number("quota");
                }
                var op = tenant.CreateSite(props);
                ctx.Load(op, o => o.IsComplete, o => o.PollingInterval);
                await ctx.RunAsync().ConfigureAwait(false);
                await Csom.WaitAsync(ctx, op, c.Cancellation).ConfigureAwait(false);
                return CommandResult.Done($"Created {props.Url}.");
            },
        };

        yield return new Command
        {
            Id = "tenant.deleted-sites",
            Title = "List deleted site collections",
            Group = "Recycle bin",
            Description = "The site collections in the tenant's recycle bin, and how many days each has left.",
            PowerShell = "Get-SPODeletedSite",
            Targets = TenantOnly,
            NeedsAdmin = true,
            ReadOnly = true,
            Glyph = "\uE81C",
            Run = async (c, _) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                var tenant = new Tenant(ctx);
                var text = new StringBuilder();
                var count = 0;
                string? next = null;
                do
                {
                    var page = tenant.GetDeletedSitePropertiesFromSharePoint(next);
                    ctx.Load(page);
                    await ctx.RunAsync().ConfigureAwait(false);
                    foreach (var d in page)
                    {
                        count++;
                        text.AppendLine($"{d.Url}  ·  deleted {d.DeletionTime:d}, {d.DaysRemaining} days left");
                    }
                    next = page.NextStartIndexFromSharePoint;
                }
                while (!string.IsNullOrEmpty(next) && count < 5000);
                return count == 0
                    ? CommandResult.Done("The recycle bin has no site collections.")
                    : CommandResult.Done($"{count:N0} deleted site collections.", text.ToString().TrimEnd());
            },
        };

        yield return new Command
        {
            Id = "tenant.restore-site",
            Title = "Restore a deleted site collection",
            Group = "Recycle bin",
            Description = "Brings a site collection back from the tenant's recycle bin.",
            PowerShell = "Restore-SPODeletedSite",
            Targets = TenantOnly,
            NeedsAdmin = true,
            After = After.ReloadSites,
            Glyph = "\uE7A7",
            Parameters = [new("url", "Address of the deleted site", ParameterKind.Url) { Placeholder = "https://contoso.sharepoint.com/sites/old-project" }],
            Run = async (c, _) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                var op = new Tenant(ctx).RestoreDeletedSite(c.Args.Text("url").TrimEnd('/'));
                ctx.Load(op, o => o.IsComplete, o => o.PollingInterval);
                await ctx.RunAsync().ConfigureAwait(false);
                await Csom.WaitAsync(ctx, op, c.Cancellation).ConfigureAwait(false);
                return CommandResult.Done($"Restored {c.Args.Text("url")}.");
            },
        };

        yield return new Command
        {
            Id = "tenant.purge-site",
            Title = "Delete a site collection for good",
            Group = "Recycle bin",
            Description = "Removes a deleted site collection from the tenant's recycle bin. It can't be restored after this.",
            PowerShell = "Remove-SPODeletedSite",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Destructive = true,
            Glyph = "\uE74D",
            Parameters = [new("url", "Address of the deleted site", ParameterKind.Url)],
            Run = async (c, _) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                var op = new Tenant(ctx).RemoveDeletedSite(c.Args.Text("url").TrimEnd('/'));
                ctx.Load(op, o => o.IsComplete, o => o.PollingInterval);
                await ctx.RunAsync().ConfigureAwait(false);
                await Csom.WaitAsync(ctx, op, c.Cancellation).ConfigureAwait(false);
                return CommandResult.Done($"Deleted {c.Args.Text("url")} for good.");
            },
        };

        yield return new Command
        {
            Id = "tenant.sharing",
            Title = "Tenant sharing",
            Group = "Sharing",
            Description = "Who SharePoint content can be shared with, anywhere in the tenant. Sites can be stricter, not looser.",
            PowerShell = "Set-SPOTenant -SharingCapability",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Glyph = "\uE72D",
            Parameters = [new("level", "Share with", ParameterKind.Choice) { Choices = SharingLevels, Default = nameof(SharingCapabilities.ExternalUserSharingOnly) }],
            Run = async (c, _) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                var tenant = new Tenant(ctx) { SharingCapability = c.Args.Enum<SharingCapabilities>("level") };
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Changed the tenant's sharing.");
            },
        };

        yield return new Command
        {
            Id = "tenant.default-link",
            Title = "Tenant's default sharing link",
            Group = "Sharing",
            Description = "The kind of link, and what it allows, that Share offers first.",
            PowerShell = "Set-SPOTenant -DefaultSharingLinkType -DefaultLinkPermission",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Glyph = "\uE71B",
            Parameters =
            [
                new("type", "Link", ParameterKind.Choice) { Choices = LinkTypes.Skip(1).ToList(), Default = nameof(SharingLinkType.Internal) },
                new("permission", "Allows", ParameterKind.Choice) { Choices = LinkPermissions.Skip(1).ToList(), Default = nameof(SharingPermissionType.View) },
            ],
            Run = async (c, _) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                var tenant = new Tenant(ctx)
                {
                    DefaultSharingLinkType = c.Args.Enum<SharingLinkType>("type"),
                    DefaultLinkPermission = c.Args.Enum<SharingPermissionType>("permission"),
                };
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Changed the tenant's default link.");
            },
        };

        yield return new Command
        {
            Id = "tenant.anyone-expiry",
            Title = "\"Anyone\" links expire",
            Group = "Sharing",
            Description = "How many days links that need no sign-in keep working, across the tenant. Sites can set fewer.",
            PowerShell = "Set-SPOTenant -RequireAnonymousLinksExpireInDays",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Glyph = "\uE916",
            Parameters = [new("days", "Expire after (days)", ParameterKind.Number) { Required = false, Min = 1, Max = 730, Hint = "Empty: they don't expire." }],
            Run = async (c, _) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                var tenant = new Tenant(ctx) { RequireAnonymousLinksExpireInDays = c.Args["days"] is null ? 0 : (int)c.Args.Number("days") };
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done(c.Args["days"] is null ? "\"Anyone\" links don't expire." : $"\"Anyone\" links expire after {c.Args.Number("days"):N0} days.");
            },
        };

        yield return new Command
        {
            Id = "tenant.guest-expiry",
            Title = "Guest access expires",
            Group = "Sharing",
            Description = "Whether guests lose access to sites after a number of days, unless an owner renews it. Sites can set their own.",
            PowerShell = "Set-SPOTenant -ExternalUserExpirationRequired -ExternalUserExpireInDays",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Glyph = "\uE916",
            Parameters = [new("days", "Expire after (days)", ParameterKind.Number) { Required = false, Min = 30, Max = 730, Hint = "Empty: guests don't expire." }],
            Run = async (c, _) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                var tenant = new Tenant(ctx) { ExternalUserExpirationRequired = c.Args["days"] is not null };
                if (c.Args["days"] is not null)
                {
                    tenant.ExternalUserExpireInDays = (int)c.Args.Number("days");
                }
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done(c.Args["days"] is null ? "Guests don't expire." : $"Guest access expires after {c.Args.Number("days"):N0} days.");
            },
        };

        yield return new Command
        {
            Id = "tenant.legacy-auth",
            Title = "Legacy sign-in",
            Group = "Settings",
            Description = "Whether apps that don't use modern authentication (old Office versions, some tools) can sign in to SharePoint.",
            PowerShell = "Set-SPOTenant -LegacyAuthProtocolsEnabled",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Glyph = "\uE8D7",
            Parameters = [Toggle("allow", "Allow legacy sign-in", false)],
            Run = async (c, _) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                var tenant = new Tenant(ctx) { LegacyAuthProtocolsEnabled = c.Args.Toggle("allow") };
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done(c.Args.Toggle("allow") ? "Legacy sign-in allowed." : "Legacy sign-in blocked.");
            },
        };

        yield return new Command
        {
            Id = "tenant.hubs",
            Title = "List hub sites",
            Group = "Hub sites",
            Description = "The tenant's hub sites, with their addresses.",
            PowerShell = "Get-SPOHubSite",
            Targets = TenantOnly,
            NeedsAdmin = true,
            ReadOnly = true,
            Glyph = "\uE71D",
            Run = async (c, _) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                var hubs = new Tenant(ctx).GetHubSitesProperties();
                ctx.Load(hubs);
                await ctx.RunAsync().ConfigureAwait(false);
                return hubs.Count == 0
                    ? CommandResult.Done("The tenant has no hub sites.")
                    : CommandResult.Done($"{hubs.Count} hub sites.", string.Join(Environment.NewLine,
                        hubs.OrderBy(h => h.Title, StringComparer.CurrentCultureIgnoreCase).Select(h => $"{h.Title}  ·  {h.SiteUrl}")));
            },
        };
    }
}
