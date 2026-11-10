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

    /// <summary>Changes tenant settings (Set-SPOTenant).</summary>
    private static async Task<CommandResult> ChangeTenantAsync(CommandContext c, Action<Tenant> change, string done)
    {
        using var ctx = await c.AdminAsync().ConfigureAwait(false);
        change(new Tenant(ctx));
        await ctx.RunAsync().ConfigureAwait(false);
        return CommandResult.Done(done);
    }

    /// <summary>Domains as typed (lines, spaces, commas) as SharePoint takes them: separated by spaces.</summary>
    internal static string Domains(string? text) =>
        string.Join(" ", (text ?? "").Split([' ', '\n', '\r', '\t', ',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));

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
            Id = "tenant.domains",
            Title = "Guests from some domains only",
            Group = "Sharing",
            Description = "Limits sharing with guests to some email domains, or blocks some, across the tenant.",
            PowerShell = "Set-SPOTenant -SharingDomainRestrictionMode -SharingAllowedDomainList -SharingBlockedDomainList",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Glyph = "\uE774",
            Parameters =
            [
                new("mode", "Guests", ParameterKind.Choice)
                {
                    Default = nameof(SharingDomainRestrictionModes.None),
                    Choices = [
                        new(nameof(SharingDomainRestrictionModes.None), "From any domain"),
                        new(nameof(SharingDomainRestrictionModes.AllowList), "Only from these domains"),
                        new(nameof(SharingDomainRestrictionModes.BlockList), "From any domain but these"),
                    ],
                },
                new("domains", "Domains", ParameterKind.LongText)
                {
                    Placeholder = "fabrikam.com",
                    Hint = "One per line, or separated by spaces.",
                    ShowIf = new ShowIf("mode", nameof(SharingDomainRestrictionModes.AllowList), nameof(SharingDomainRestrictionModes.BlockList)),
                },
            ],
            Run = (c, _) =>
            {
                var mode = c.Args.Enum<SharingDomainRestrictionModes>("mode");
                var domains = Domains(c.Args["domains"]);
                return ChangeTenantAsync(c, t =>
                {
                    t.SharingDomainRestrictionMode = mode;
                    if (mode == SharingDomainRestrictionModes.AllowList)
                    {
                        t.SharingAllowedDomainList = domains;
                    }
                    else if (mode == SharingDomainRestrictionModes.BlockList)
                    {
                        t.SharingBlockedDomainList = domains;
                    }
                }, mode switch
                {
                    SharingDomainRestrictionModes.AllowList => $"Guests only from {domains}.",
                    SharingDomainRestrictionModes.BlockList => $"Guests from anywhere but {domains}.",
                    _ => "Guests from any domain.",
                });
            },
        };

        yield return new Command
        {
            Id = "tenant.reshare",
            Title = "Guests reshare",
            Group = "Sharing",
            Description = "Whether guests can share files and folders they don't own.",
            PowerShell = "Set-SPOTenant -PreventExternalUsersFromResharing",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Glyph = "\uE72D",
            Parameters = [Toggle("allow", "Guests can share what they don't own", true)],
            Run = (c, _) => ChangeTenantAsync(c, t => t.PreventExternalUsersFromResharing = !c.Args.Toggle("allow"),
                c.Args.Toggle("allow") ? "Guests can reshare." : "Guests can't reshare."),
        };

        yield return new Command
        {
            Id = "tenant.everyone",
            Title = "\"Everyone\" in the people picker",
            Group = "Sharing",
            Description = "Whether people can pick Everyone, or Everyone except external users, when they share or give access.",
            PowerShell = "Set-SPOTenant -ShowEveryoneClaim -ShowEveryoneExceptExternalUsersClaim",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Glyph = "\uE716",
            Parameters =
            [
                Toggle("everyone", "Offer \"Everyone\" (guests included)", false),
                Toggle("except", "Offer \"Everyone except external users\"", true),
            ],
            Run = (c, _) => ChangeTenantAsync(c, t =>
            {
                t.ShowEveryoneClaim = c.Args.Toggle("everyone");
                t.ShowEveryoneExceptExternalUsersClaim = c.Args.Toggle("except");
            }, "Changed the people picker."),
        };

        yield return new Command
        {
            Id = "tenant.notifications",
            Title = "Sharing emails to owners",
            Group = "Sharing",
            Description = "Which sharing events owners of OneDrive and sites are told about by email.",
            PowerShell = "Set-SPOTenant -NotifyOwnersWhenItemsReshared -NotifyOwnersWhenInvitationsAccepted -OwnerAnonymousNotification",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Glyph = "\uE715",
            Parameters =
            [
                Toggle("reshared", "When someone reshares their files", true),
                Toggle("accepted", "When a guest accepts an invitation", true),
                Toggle("anyone", "When someone makes an \"Anyone\" link", true),
            ],
            Run = (c, _) => ChangeTenantAsync(c, t =>
            {
                t.NotifyOwnersWhenItemsReshared = c.Args.Toggle("reshared");
                t.NotifyOwnersWhenInvitationsAccepted = c.Args.Toggle("accepted");
                t.OwnerAnonymousNotification = c.Args.Toggle("anyone");
            }, "Changed the emails to owners."),
        };

        yield return new Command
        {
            Id = "tenant.site-creation",
            Title = "People create sites",
            Group = "Settings",
            Description = "Whether people can create sites themselves, from SharePoint's start page and OneDrive.",
            PowerShell = "Set-SPOTenant -SelfServiceSiteCreationDisabled",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Glyph = "\uE710",
            Parameters = [Toggle("allow", "Allow people to create sites", true)],
            Run = (c, _) => ChangeTenantAsync(c, t => t.SelfServiceSiteCreationDisabled = !c.Args.Toggle("allow"),
                c.Args.Toggle("allow") ? "People can create sites." : "Only administrators create sites."),
        };

        yield return new Command
        {
            Id = "tenant.unmanaged",
            Title = "Access from unmanaged devices",
            Group = "Settings",
            Description = "Full access, web-only access (no downloading, printing or syncing), or none, for devices Conditional Access doesn't manage, across the tenant.",
            PowerShell = "Set-SPOTenant -ConditionalAccessPolicy",
            Targets = TenantOnly,
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
            Run = (c, _) => ChangeTenantAsync(c, t => t.ConditionalAccessPolicy = c.Args.Enum<SPOConditionalAccessPolicyType>("policy"),
                "Changed access from unmanaged devices."),
        };

        yield return new Command
        {
            Id = "tenant.home-site",
            Title = "Home site",
            Group = "Settings",
            Description = "The communication site that's the intranet's front page (the SharePoint app bar's home). Empty takes it away.",
            PowerShell = "Set-SPOHomeSite, Remove-SPOHomeSite",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Glyph = "\uE80F",
            Parameters = [new("url", "Site address", ParameterKind.Url) { Required = false, Placeholder = "https://contoso.sharepoint.com/sites/intranet" }],
            Run = async (c, _) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                var tenant = new Tenant(ctx);
                if (c.Args["url"] is { } url)
                {
                    tenant.SetSPHSite(url.TrimEnd('/'));
                }
                else
                {
                    tenant.RemoveSPHSite();
                }
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done(c.Args["url"] is { } u ? $"{u} is the home site." : "No home site now.");
            },
        };

        yield return new Command
        {
            Id = "tenant.onedrive-quota",
            Title = "OneDrive storage",
            Group = "Settings",
            Description = "How much each new OneDrive may store. Existing ones keep theirs.",
            PowerShell = "Set-SPOTenant -OneDriveStorageQuota",
            Targets = TenantOnly,
            NeedsAdmin = true,
            Glyph = "\uEDA2",
            Parameters = [new("mb", "Limit (MB)", ParameterKind.Number) { Default = "1048576", Min = 1024, Max = 25L * 1024 * 1024, Hint = "1048576 is 1 TB; 5242880 is 5 TB." }],
            Run = (c, _) => ChangeTenantAsync(c, t => t.OneDriveStorageQuota = c.Args.Number("mb"), $"New OneDrives get {c.Args.Number("mb"):N0} MB."),
        };

        yield return new Command
        {
            Id = "tenant.site-designs",
            Title = "List site designs",
            Group = "Details",
            Description = "The tenant's site designs (site templates), with their IDs for Apply a site design.",
            PowerShell = "Get-SPOSiteDesign",
            Targets = TenantOnly,
            NeedsAdmin = true,
            ReadOnly = true,
            Glyph = "\uE771",
            Run = async (c, _) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                var designs = new Tenant(ctx).GetSiteDesigns();
                ctx.Load(designs);
                await ctx.RunAsync().ConfigureAwait(false);
                return designs.Count == 0
                    ? CommandResult.Done("The tenant has no site designs of its own.")
                    : CommandResult.Done($"{designs.Count} site designs.", string.Join(Environment.NewLine, designs.ToList()
                        .OrderBy(d => d.Title, StringComparer.CurrentCultureIgnoreCase)
                        .Select(d => $"{d.Title}  ·  {d.Id}  ·  {(d.WebTemplate == "68" ? "communication site" : d.WebTemplate == "64" ? "team site" : d.WebTemplate)}")));
            },
        };

        yield return new Command
        {
            Id = "tenant.themes",
            Title = "List themes",
            Group = "Details",
            Description = "The themes added to the tenant, by the names Theme takes.",
            PowerShell = "Get-SPOTheme",
            Targets = TenantOnly,
            NeedsAdmin = true,
            ReadOnly = true,
            Glyph = "\uE790",
            Run = async (c, _) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                var themes = new Tenant(ctx).GetAllTenantThemes();
                ctx.Load(themes);
                await ctx.RunAsync().ConfigureAwait(false);
                return themes.Count == 0
                    ? CommandResult.Done("The tenant has no themes of its own.")
                    : CommandResult.Done($"{themes.Count} themes.", string.Join(Environment.NewLine, themes.ToList().Select(t => t.Name).Order(StringComparer.CurrentCultureIgnoreCase)));
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
                        hubs.ToList().OrderBy(h => h.Title, StringComparer.CurrentCultureIgnoreCase).Select(h => $"{h.Title}  ·  {h.SiteUrl}")));
            },
        };
    }
}
