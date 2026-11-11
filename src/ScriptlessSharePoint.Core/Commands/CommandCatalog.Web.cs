using System.Globalization;
using System.Text;

using Microsoft.SharePoint.Client;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>Sites as webs: a site collection's root web and its subsites (Set-PnPWeb and friends).</summary>
public static partial class CommandCatalog
{
    private static readonly IReadOnlyList<Choice> ListTemplates =
    [
        new(((int)ListTemplateType.GenericList).ToString(CultureInfo.InvariantCulture), "Custom list"),
        new(((int)ListTemplateType.Tasks).ToString(CultureInfo.InvariantCulture), "Tasks"),
        new(((int)ListTemplateType.IssueTracking).ToString(CultureInfo.InvariantCulture), "Issue tracking"),
        new(((int)ListTemplateType.Contacts).ToString(CultureInfo.InvariantCulture), "Contacts"),
        new(((int)ListTemplateType.Links).ToString(CultureInfo.InvariantCulture), "Links"),
        new(((int)ListTemplateType.Events).ToString(CultureInfo.InvariantCulture), "Calendar"),
        new(((int)ListTemplateType.Announcements).ToString(CultureInfo.InvariantCulture), "Announcements"),
    ];

    private static CommandParameter NavigationPlace() => new("where", "In", ParameterKind.Choice)
    {
        Default = "side",
        Choices = [new("side", "The site's navigation"), new("top", "The top navigation (communication sites)")],
    };

    private static NavigationNodeCollection Navigation(ClientContext ctx, string where) =>
        where == "top" ? ctx.Web.Navigation.TopNavigationBar : ctx.Web.Navigation.QuickLaunch;

    private static async Task<CommandResult> ChangeWebAsync(CommandContext c, SpNode n, Action<Microsoft.SharePoint.Client.Web> change, string done)
    {
        using var ctx = await c.WebAsync(n).ConfigureAwait(false);
        change(ctx.Web);
        ctx.Web.Update();
        await ctx.RunAsync().ConfigureAwait(false);
        return CommandResult.Done(done);
    }

    /// <summary>"SitePages/Home.aspx" from that, from "/sites/hr/SitePages/Home.aspx", or from the full address.</summary>
    internal static string WithinWeb(string webUrl, string page)
    {
        var path = page.Trim();
        if (Uri.TryCreate(path, UriKind.Absolute, out var uri) && uri.Scheme.StartsWith("http", StringComparison.OrdinalIgnoreCase))
        {
            path = Uri.UnescapeDataString(uri.AbsolutePath);
        }
        var web = Uri.TryCreate(webUrl, UriKind.Absolute, out var w) ? Uri.UnescapeDataString(w.AbsolutePath).TrimEnd('/') : "";
        if (web.Length > 0 && path.StartsWith(web + "/", StringComparison.OrdinalIgnoreCase))
        {
            path = path[(web.Length + 1)..];
        }
        return path.TrimStart('/');
    }

    private static string Shorten(string text, int max) => text.Length <= max ? text : text[..max] + "…";

    private static IEnumerable<Command> Web()
    {
        yield return new Command
        {
            Id = "web.title",
            Title = "Rename",
            Group = "Settings",
            Description = "Changes each site's title, and its description when one is given.",
            PowerShell = "Set-PnPWeb -Title -Description",
            Targets = Webs,
            After = After.Update,
            Glyph = "\uE8AC",
            Parameters =
            [
                new("title", "Title", ParameterKind.Text),
                new("description", "Description", ParameterKind.LongText) { Required = false, Hint = "Empty: the description stays as it is." },
            ],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                ctx.Web.Title = c.Args.Text("title");
                if (c.Args["description"] is { } d)
                {
                    ctx.Web.Description = d;
                }
                ctx.Web.Update();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Renamed.") with { Updated = n with { Title = c.Args.Text("title") } };
            },
        };

        yield return new Command
        {
            Id = "web.search",
            Title = "Show in search results",
            Group = "Search",
            Description = "Whether the site's content shows up in search. Where custom scripts are blocked, they're allowed for the moment (needs a SharePoint administrator).",
            PowerShell = "Set-PnPWeb -NoCrawl",
            Targets = Webs,
            Glyph = "\uE721",
            Parameters = [Toggle("show", "Show in search", true)],
            Run = async (c, n) =>
            {
                // SharePoint refuses NoCrawl where custom scripts are blocked, as on every modern site
                var allowed = await c.WithScriptsAsync(n, async () =>
                {
                    using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                    ctx.Web.NoCrawl = !c.Args.Toggle("show");
                    ctx.Web.Update();
                    await ctx.RunAsync().ConfigureAwait(false);
                }).ConfigureAwait(false);
                return CommandResult.Done(Csom.ScriptsNote(c.Args.Toggle("show") ? "Shown in search." : "Left out of search.", allowed));
            },
        };

        yield return new Command
        {
            Id = "web.reindex",
            Title = "Reindex",
            Group = "Search",
            Description = "Asks search to crawl the whole site again at its next crawl, e.g. after changing its search settings. Where custom scripts are blocked, they're allowed for the moment (needs a SharePoint administrator).",
            PowerShell = "Request-PnPReindexWeb",
            Targets = Webs,
            Glyph = "\uE72C",
            Run = async (c, n) =>
            {
                var allowed = await c.WithScriptsAsync(n, async () =>
                {
                    using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                    var props = ctx.Web.AllProperties;
                    ctx.Load(props);
                    await ctx.RunAsync().ConfigureAwait(false);
                    var version = props.FieldValues.TryGetValue("vti_searchversion", out var v) && int.TryParse(v?.ToString(), out var i) ? i : 0;
                    props["vti_searchversion"] = version + 1;
                    ctx.Web.Update();
                    await ctx.RunAsync().ConfigureAwait(false);
                }).ConfigureAwait(false);
                return CommandResult.Done(Csom.ScriptsNote("Search will crawl it again.", allowed));
            },
        };

        yield return new Command
        {
            Id = "web.members-share",
            Title = "Members can share",
            Group = "Sharing",
            Description = "Whether the site's members may share files, folders and the site itself, or only its owners.",
            PowerShell = "Set-PnPWeb -MembersCanShare",
            Targets = Webs,
            Glyph = "\uE72D",
            Parameters = [Toggle("allow", "Members can share", true)],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                ctx.Web.MembersCanShare = c.Args.Toggle("allow");
                ctx.Web.Update();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done(c.Args.Toggle("allow") ? "Members can share." : "Only owners can share.");
            },
        };

        yield return new Command
        {
            Id = "web.property",
            Title = "Set a property bag value",
            Group = "Settings",
            Description = "Stores a value in each site's property bag. Where custom scripts are blocked, they're allowed for the moment (needs a SharePoint administrator).",
            PowerShell = "Set-PnPPropertyBagValue",
            Targets = Webs,
            Glyph = "\uE8EC",
            Parameters =
            [
                new("key", "Key", ParameterKind.Text),
                new("value", "Value", ParameterKind.Text) { Required = false, Hint = "Empty stores an empty value." },
            ],
            Run = async (c, n) =>
            {
                var allowed = await c.WithScriptsAsync(n, async () =>
                {
                    using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                    ctx.Web.AllProperties[c.Args.Text("key")] = c.Args["value"] ?? "";
                    ctx.Web.Update();
                    await ctx.RunAsync().ConfigureAwait(false);
                }).ConfigureAwait(false);
                return CommandResult.Done(Csom.ScriptsNote($"Set {c.Args.Text("key")}.", allowed));
            },
        };

        yield return new Command
        {
            Id = "web.properties",
            Title = "Show the property bag",
            Group = "Details",
            Description = "The keys and values in each site's property bag.",
            PowerShell = "Get-PnPPropertyBag",
            Targets = Webs,
            ReadOnly = true,
            Glyph = "\uE8EC",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var props = ctx.Web.AllProperties;
                ctx.Load(props);
                await ctx.RunAsync().ConfigureAwait(false);
                var rows = props.FieldValues.OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase)
                    .Select(p => new[] { p.Key, Shorten(Convert.ToString(p.Value, CultureInfo.InvariantCulture) ?? "", 200) })
                    .ToList();
                return CommandResult.Report($"{rows.Count:N0} values.", ReportTable.Of(["Key", "Value"], rows));
            },
        };

        yield return new Command
        {
            Id = "web.property-remove",
            Title = "Remove a property bag value",
            Group = "Settings",
            Description = "Removes a key from each site's property bag. Where custom scripts are blocked, they're allowed for the moment (needs a SharePoint administrator).",
            PowerShell = "Remove-PnPPropertyBagValue",
            Targets = Webs,
            Glyph = "\uE8EC",
            Parameters = [new("key", "Key", ParameterKind.Text)],
            Run = async (c, n) =>
            {
                var key = c.Args.Text("key");
                var found = true;
                var allowed = await c.WithScriptsAsync(n, async () =>
                {
                    using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                    var props = ctx.Web.AllProperties;
                    ctx.Load(props);
                    await ctx.RunAsync().ConfigureAwait(false);
                    found = props.FieldValues.ContainsKey(key);
                    if (found)
                    {
                        props[key] = null;
                        props.FieldValues.Remove(key);
                        ctx.Web.Update();
                        await ctx.RunAsync().ConfigureAwait(false);
                    }
                }).ConfigureAwait(false);
                return found
                    ? CommandResult.Done(Csom.ScriptsNote($"Removed {key}.", allowed))
                    : CommandResult.Skipped($"It has no {key}.");
            },
        };

        yield return new Command
        {
            Id = "web.home-page",
            Title = "Home page",
            Group = "Settings",
            Description = "Which page each site opens on.",
            PowerShell = "Set-PnPHomePage",
            Targets = Webs,
            Glyph = "\uE80F",
            Parameters = [new("page", "Page", ParameterKind.Text) { Default = "SitePages/Home.aspx", Hint = "Its address within the site, as SitePages/Home.aspx (a full address works too)." }],
            Run = async (c, n) =>
            {
                var page = WithinWeb(n.WebUrl, c.Args.Text("page"));
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                ctx.Web.RootFolder.WelcomePage = page;
                ctx.Web.RootFolder.Update();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Opens on {page}.");
            },
        };

        yield return new Command
        {
            Id = "web.logo",
            Title = "Site logo",
            Group = "Settings",
            Description = "The logo in each site's header: an image's address, in the site or elsewhere in the tenant. Empty puts back the default.",
            PowerShell = "Set-PnPWeb -SiteLogoUrl",
            Targets = Webs,
            Glyph = "\uEB9F",
            Parameters = [new("logo", "Image address", ParameterKind.Text) { Required = false, Placeholder = "/sites/hr/SiteAssets/logo.png" }],
            Run = (c, n) => ChangeWebAsync(c, n, w => w.SiteLogoUrl = c.Args["logo"]?.Trim() ?? "", c.Args["logo"] is null ? "The default logo." : "Logo changed."),
        };

        yield return new Command
        {
            Id = "web.header",
            Title = "Header",
            Group = "Settings",
            Description = "The layout and background of each site's header, and whether it shows the site's title.",
            PowerShell = "Set-PnPWeb -HeaderLayout -HeaderEmphasis -HideTitleInHeader",
            Targets = Webs,
            Glyph = "\uE7C3",
            Parameters =
            [
                new("layout", "Layout", ParameterKind.Choice)
                {
                    Default = nameof(HeaderLayoutType.Standard),
                    Choices =
                    [
                        new(nameof(HeaderLayoutType.Standard), "Standard"),
                        new(nameof(HeaderLayoutType.Compact), "Compact"),
                        new(nameof(HeaderLayoutType.Minimal), "Minimal"),
                        new(nameof(HeaderLayoutType.Extended), "Extended"),
                    ],
                },
                new("emphasis", "Background", ParameterKind.Choice)
                {
                    Default = nameof(SPVariantThemeType.None),
                    Choices =
                    [
                        new(nameof(SPVariantThemeType.None), "None"),
                        new(nameof(SPVariantThemeType.Neutral), "Neutral"),
                        new(nameof(SPVariantThemeType.Soft), "Soft"),
                        new(nameof(SPVariantThemeType.Strong), "Strong"),
                    ],
                },
                Toggle("title", "Show the site's title", true),
            ],
            Run = (c, n) => ChangeWebAsync(c, n, w =>
            {
                w.HeaderLayout = c.Args.Enum<HeaderLayoutType>("layout");
                w.HeaderEmphasis = c.Args.Enum<SPVariantThemeType>("emphasis");
                w.HideTitleInHeader = !c.Args.Toggle("title");
            }, "Header changed."),
        };

        yield return new Command
        {
            Id = "web.navigation",
            Title = "Navigation",
            Group = "Settings",
            Description = "Whether each site shows its navigation, where, and how.",
            PowerShell = "Set-PnPWeb -QuickLaunchEnabled -HorizontalQuickLaunch -MegaMenuEnabled",
            Targets = Webs,
            Glyph = "\uE700",
            Parameters =
            [
                Toggle("show", "Show the site's navigation", true),
                new("place", "Place", ParameterKind.Choice)
                {
                    Default = "side",
                    Choices = [new("side", "Down the side"), new("top", "Across the top")],
                    ShowIf = new ShowIf("show", "true"),
                },
                Toggle("mega", "Mega menu", false, "For navigation across the top: links in wide panels.") with { ShowIf = new ShowIf("show", "true") },
            ],
            Run = (c, n) => ChangeWebAsync(c, n, w =>
            {
                w.QuickLaunchEnabled = c.Args.Toggle("show");
                if (c.Args.Toggle("show"))
                {
                    w.HorizontalQuickLaunch = c.Args.Text("place") == "top";
                    w.MegaMenuEnabled = c.Args.Toggle("mega");
                }
            }, c.Args.Toggle("show") ? "Navigation changed." : "Navigation hidden."),
        };

        yield return new Command
        {
            Id = "web.footer",
            Title = "Footer",
            Group = "Settings",
            Description = "Whether each site shows its footer (communication sites).",
            PowerShell = "Set-PnPFooter -Enabled",
            Targets = Webs,
            Glyph = "\uE8A0",
            Parameters = [Toggle("show", "Show the footer", true)],
            Run = (c, n) => ChangeWebAsync(c, n, w => w.FooterEnabled = c.Args.Toggle("show"), c.Args.Toggle("show") ? "Footer shown." : "Footer hidden."),
        };

        yield return new Command
        {
            Id = "web.nav-add",
            Title = "Add a navigation link",
            Group = "Settings",
            Description = "Adds a link at the end of each site's navigation.",
            PowerShell = "Add-PnPNavigationNode",
            Targets = Webs,
            Glyph = "\uE71B",
            Parameters =
            [
                new("title", "Title", ParameterKind.Text),
                new("url", "Address", ParameterKind.Text) { Placeholder = "https://contoso.sharepoint.com/sites/hr", Hint = "A full address, or one in the tenant as /sites/hr/SitePages/Policies.aspx." },
                NavigationPlace(),
            ],
            Run = async (c, n) =>
            {
                var url = c.Args.Text("url").Trim();
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                Navigation(ctx, c.Args.Text("where")).Add(new NavigationNodeCreationInformation
                {
                    Title = c.Args.Text("title"),
                    Url = url,
                    AsLastNode = true,
                    IsExternal = Uri.TryCreate(url, UriKind.Absolute, out _),
                });
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Added {c.Args.Text("title")}.");
            },
        };

        yield return new Command
        {
            Id = "web.nav-remove",
            Title = "Remove a navigation link",
            Group = "Settings",
            Description = "Removes the links with this title from each site's navigation, at the top level and one below.",
            PowerShell = "Remove-PnPNavigationNode",
            Targets = Webs,
            Glyph = "\uE711",
            Parameters = [new("title", "Title", ParameterKind.Text), NavigationPlace()],
            Run = async (c, n) =>
            {
                var title = c.Args.Text("title").Trim();
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var nodes = Navigation(ctx, c.Args.Text("where"));
                ctx.Load(nodes, ns => ns.Include(x => x.Title, x => x.Children.Include(y => y.Title)));
                await ctx.RunAsync().ConfigureAwait(false);
                var matches = nodes.ToList().Concat(nodes.ToList().SelectMany(x => x.Children.ToList()))
                    .Where(x => string.Equals(x.Title, title, StringComparison.CurrentCultureIgnoreCase))
                    .ToList();
                if (matches.Count == 0)
                {
                    return CommandResult.Skipped($"No link called {title}.");
                }
                foreach (var node in matches)
                {
                    node.DeleteObject();
                }
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done(matches.Count == 1 ? $"Removed {title}." : $"Removed {matches.Count} links called {title}.");
            },
        };

        yield return new Command
        {
            Id = "web.regional",
            Title = "Time zone and locale",
            Group = "Settings",
            Description = "The time zone each site shows times in, and its locale (how dates and numbers are written).",
            PowerShell = "Set-PnPWeb (RegionalSettings.TimeZone, RegionalSettings.LocaleId)",
            Targets = Webs,
            Glyph = "\uEC92",
            Parameters =
            [
                new("zone", "Time zone", ParameterKind.Text)
                {
                    Required = false,
                    Placeholder = "Copenhagen",
                    Hint = "A city or offset from SharePoint's list of time zones, as Copenhagen or UTC+01:00. Empty leaves it.",
                },
                new("locale", "Locale (LCID)", ParameterKind.Number)
                {
                    Required = false,
                    Min = 1025,
                    Max = 58378,
                    Hint = "1033 English (United States), 2057 English (United Kingdom), 1030 Danish, 1031 German. Empty leaves it.",
                },
            ],
            Run = async (c, n) =>
            {
                if (c.Args["zone"] is null && c.Args["locale"] is null)
                {
                    return CommandResult.Skipped("Nothing to change: give a time zone or a locale.");
                }
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var settings = ctx.Web.RegionalSettings;
                var changed = new List<string>();
                if (c.Args["zone"] is { } zone)
                {
                    var zones = settings.TimeZones;
                    ctx.Load(zones, z => z.Include(x => x.Id, x => x.Description));
                    await ctx.RunAsync().ConfigureAwait(false);
                    var term = zone.Trim();
                    var matches = zones.ToList().Where(z => z.Description.Contains(term, StringComparison.OrdinalIgnoreCase)).ToList();
                    if (matches.Count != 1)
                    {
                        return CommandResult.Failed(matches.Count == 0
                            ? $"No time zone matches “{term}”."
                            : $"“{term}” matches {matches.Count} time zones: {string.Join("; ", matches.Take(4).Select(z => z.Description))}{(matches.Count > 4 ? "…" : "")}");
                    }
                    settings.TimeZone = matches[0];
                    changed.Add(matches[0].Description);
                }
                if (c.Args["locale"] is not null)
                {
                    settings.LocaleId = (uint)c.Args.Number("locale");
                    changed.Add($"locale {c.Args.Number("locale")}");
                }
                settings.Update();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Set {string.Join(" and ", changed)}.");
            },
        };

        yield return new Command
        {
            Id = "web.access-requests",
            Title = "Access requests",
            Group = "Sharing",
            Description = "Whether people without access can ask for it, and who gets their requests. Only for sites with their own permissions.",
            PowerShell = "Set-PnPRequestAccessEmails",
            Targets = Webs,
            Glyph = "\uE715",
            Parameters =
            [
                new("to", "Requests go to", ParameterKind.Choice)
                {
                    Default = "owners",
                    Choices = [new("owners", "The site's owners"), new("email", "An email address"), new("off", "Nobody: turn requests off")],
                },
                new("email", "Email address", ParameterKind.Text) { Placeholder = "helpdesk@contoso.com", ShowIf = new ShowIf("to", "email") },
            ],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var to = c.Args.Text("to");
                if (to == "owners")
                {
                    ctx.Web.SetUseAccessRequestDefaultAndUpdate(true);
                }
                else
                {
                    ctx.Web.SetUseAccessRequestDefaultAndUpdate(false);
                    ctx.Web.RequestAccessEmail = to == "email" ? c.Args.Text("email").Trim() : "";
                    ctx.Web.Update();
                }
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done(to switch
                {
                    "owners" => "Requests go to the owners.",
                    "email" => $"Requests go to {c.Args.Text("email").Trim()}.",
                    _ => "Access requests turned off.",
                });
            },
        };

        yield return new Command
        {
            Id = "web.features",
            Title = "List active features",
            Group = "Apps and features",
            Description = "The features turned on in each site and in its site collection.",
            PowerShell = "Get-PnPFeature",
            Targets = Webs,
            ReadOnly = true,
            Glyph = "\uE74C",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var web = ctx.Web.Features;
                var site = ctx.Site.Features;
                ctx.Load(web, fs => fs.Include(f => f.DefinitionId, f => f.DisplayName));
                ctx.Load(site, fs => fs.Include(f => f.DefinitionId, f => f.DisplayName));
                await ctx.RunAsync().ConfigureAwait(false);
                static string Lines(FeatureCollection fs) => string.Join(Environment.NewLine, fs.ToList()
                    .OrderBy(f => f.DisplayName, StringComparer.CurrentCultureIgnoreCase)
                    .Select(f => $"  {(string.IsNullOrEmpty(f.DisplayName) ? "(no name)" : f.DisplayName)}  ·  {f.DefinitionId}"));
                return CommandResult.Done($"{web.Count} on the site, {site.Count} on the site collection.",
                    $"The site:{Environment.NewLine}{Lines(web)}{Environment.NewLine}The site collection:{Environment.NewLine}{Lines(site)}");
            },
        };

        yield return new Command
        {
            Id = "web.group-new",
            Title = "New group",
            Group = "Groups",
            Description = "Creates a SharePoint group in each site's site collection and, unless the site inherits its permissions, gives it a permission level there.",
            PowerShell = "New-PnPGroup",
            Targets = Webs,
            Glyph = "\uE902",
            Parameters =
            [
                new("name", "Name", ParameterKind.Text),
                new("description", "Description", ParameterKind.LongText) { Required = false },
                new("role", "Permission level on the site", ParameterKind.Choice) { Choices = [new("none", "None for now"), .. Csom.Roles], Default = "Reader" },
            ],
            Run = async (c, n) =>
            {
                var name = c.Args.Text("name").Trim();
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                ctx.Load(ctx.Web, w => w.HasUniqueRoleAssignments);
                var group = ctx.Web.SiteGroups.Add(new GroupCreationInformation { Title = name, Description = c.Args["description"] ?? "" });
                await ctx.RunAsync().ConfigureAwait(false);
                var role = c.Args.Text("role");
                if (role == "none")
                {
                    return CommandResult.Done($"Created {name}.");
                }
                if (!ctx.Web.HasUniqueRoleAssignments)
                {
                    return CommandResult.Done($"Created {name}. The site inherits its permissions: give the group access where they're set.");
                }
                ctx.Web.RoleAssignments.Add(group, new RoleDefinitionBindingCollection(ctx) { ctx.Web.RoleDefinitions.GetByType(Csom.Role(role)) });
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Created {name} with {Csom.Roles.First(r => r.Value == role).Label.ToLowerInvariant()}.");
            },
        };

        yield return new Command
        {
            Id = "web.theme",
            Title = "Theme",
            Group = "Settings",
            Description = "Applies one of the tenant's own themes to each site.",
            PowerShell = "Set-SPOWebTheme",
            Targets = Webs,
            NeedsAdmin = true,
            Glyph = "\uE790",
            Parameters = [new("theme", "Theme", ParameterKind.Text) { Hint = "Its name, as List themes (on the tenant) shows it." }],
            Run = async (c, n) =>
            {
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                new Microsoft.Online.SharePoint.TenantAdministration.Tenant(ctx).SetWebTheme(c.Args.Text("theme").Trim(), n.WebUrl);
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Applied {c.Args.Text("theme").Trim()}.");
            },
        };

        yield return new Command
        {
            Id = "web.site-design",
            Title = "Apply a site design",
            Group = "Apps and features",
            Description = "Runs one of the tenant's site designs (site templates) on each site. SharePoint does it in the background, in a few minutes.",
            PowerShell = "Add-SPOSiteDesignTask",
            Targets = Webs,
            NeedsAdmin = true,
            Glyph = "\uE771",
            Parameters = [new("design", "Site design ID", ParameterKind.Text) { Placeholder = "a1b2c3d4-…", Hint = "List site designs (on the tenant) shows the IDs." }],
            Run = async (c, n) =>
            {
                if (!Guid.TryParse(c.Args.Text("design").Trim(), out var id))
                {
                    return CommandResult.Failed("The site design ID isn't a GUID.");
                }
                using var ctx = await c.AdminAsync().ConfigureAwait(false);
                var task = Microsoft.Online.SharePoint.TenantAdministration.Tenant.AddSiteDesignTask(ctx, n.WebUrl, id);
                ctx.Load(task);
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Queued: SharePoint applies it within a few minutes.");
            },
        };

        yield return new Command
        {
            Id = "web.offline",
            Title = "Offline and sync",
            Group = "Settings",
            Description = "Whether the site's libraries can be synced with OneDrive and used offline (Search and offline availability).",
            Targets = Webs,
            Glyph = "\uE895",
            Parameters = [Toggle("allow", "Allow sync and offline use", true)],
            Run = (c, n) => ChangeWebAsync(c, n, w => w.ExcludeFromOfflineClient = !c.Args.Toggle("allow"),
                c.Args.Toggle("allow") ? "Sync allowed." : "Sync turned off."),
        };

        yield return new Command
        {
            Id = "web.group-delete",
            Title = "Delete a group",
            Group = "Groups",
            Description = "Deletes a SharePoint group from each site's site collection. Its members lose what the group gave them.",
            PowerShell = "Remove-PnPGroup",
            Targets = Webs,
            Destructive = true,
            Glyph = "\uE74D",
            Parameters = [new("name", "Group", ParameterKind.Text) { Hint = "Its name, as Site permissions shows it." }],
            Run = async (c, n) =>
            {
                var name = c.Args.Text("name").Trim();
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var group = ctx.Web.SiteGroups.GetByName(name);
                ctx.Load(group, g => g.Id);
                try
                {
                    await ctx.RunAsync().ConfigureAwait(false);
                }
                catch (ServerException)
                {
                    return CommandResult.Skipped($"It has no group called {name}.");
                }
                ctx.Web.SiteGroups.RemoveById(group.Id);
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Deleted {name}.");
            },
        };

        // ---- create ----

        yield return new Command
        {
            Id = "web.new-subsite",
            Title = "New subsite",
            Group = "Create",
            Description = "Creates a subsite in each site. (Microsoft recommends hub sites over subsites for new work.)",
            PowerShell = "New-PnPWeb",
            Targets = Webs,
            After = After.ReloadChildren,
            Glyph = "\uE710",
            Parameters =
            [
                new("title", "Title", ParameterKind.Text),
                new("url", "Address", ParameterKind.Text) { Placeholder = "projects", Hint = "The last part of its address, after the parent site's." },
                new("template", "Kind of site", ParameterKind.Choice)
                {
                    Default = "STS#3",
                    Choices = [new("STS#3", "Team site"), new("SITEPAGEPUBLISHING#0", "Communication site"), new("STS#0", "Classic team site")],
                },
                Toggle("inherit", "Use the parent site's permissions", true),
                new("lcid", "Language (LCID)", ParameterKind.Number) { Default = "1033", Min = 1025, Max = 1106 },
            ],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var leaf = c.Args.Text("url").Trim('/');
                var web = ctx.Web.Webs.Add(new WebCreationInformation
                {
                    Title = c.Args.Text("title"),
                    Url = leaf,
                    WebTemplate = c.Args.Text("template"),
                    UseSamePermissionsAsParentSite = c.Args.Toggle("inherit"),
                    Language = (int)c.Args.Number("lcid"),
                });
                ctx.Load(web, w => w.Url);
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Created {web.Url}.");
            },
        };

        yield return new Command
        {
            Id = "web.new-library",
            Title = "New document library",
            Group = "Create",
            Description = "Creates a document library in each site.",
            PowerShell = "New-PnPList -Template DocumentLibrary",
            Targets = Webs,
            After = After.ReloadChildren,
            Glyph = "\uE8F1",
            Parameters =
            [
                new("title", "Name", ParameterKind.Text),
                new("description", "Description", ParameterKind.Text) { Required = false },
                Toggle("nav", "Show on the site's navigation", true),
            ],
            Run = (c, n) => NewListAsync(c, n, (int)ListTemplateType.DocumentLibrary),
        };

        yield return new Command
        {
            Id = "web.new-list",
            Title = "New list",
            Group = "Create",
            Description = "Creates a list in each site.",
            PowerShell = "New-PnPList",
            Targets = Webs,
            After = After.ReloadChildren,
            Glyph = "\uE8FD",
            Parameters =
            [
                new("title", "Name", ParameterKind.Text),
                new("template", "Kind of list", ParameterKind.Choice) { Choices = ListTemplates, Default = ListTemplates[0].Value },
                new("description", "Description", ParameterKind.Text) { Required = false },
                Toggle("nav", "Show on the site's navigation", true),
            ],
            Run = (c, n) => NewListAsync(c, n, int.Parse(c.Args.Text("template"), CultureInfo.InvariantCulture)),
        };

        // ---- groups ----

        yield return new Command
        {
            Id = "web.group-add",
            Title = "Add to the site's owners, members or visitors",
            Group = "Groups",
            Description = "Adds a person or group to one of each site's own SharePoint groups.",
            PowerShell = "Add-PnPGroupMember",
            Targets = Webs,
            Glyph = "\uE8FA",
            Parameters =
            [
                new("group", "Group", ParameterKind.Choice) { Choices = Csom.SiteGroups, Default = "Members" },
                Person(),
            ],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var group = Csom.SiteGroup(ctx.Web, c.Args.Text("group"));
                group.Users.AddUser(Csom.EnsureUser(ctx, c.Args.Text("who")));
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Added to the {c.Args.Text("group").ToLowerInvariant()}.");
            },
        };

        yield return new Command
        {
            Id = "web.group-remove",
            Title = "Remove from the site's owners, members or visitors",
            Group = "Groups",
            Description = "Takes a person or group out of one of each site's own SharePoint groups.",
            PowerShell = "Remove-PnPGroupMember",
            Targets = Webs,
            Glyph = "\uE8AB",
            Parameters =
            [
                new("group", "Group", ParameterKind.Choice) { Choices = Csom.SiteGroups, Default = "Members" },
                Person(),
            ],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var group = Csom.SiteGroup(ctx.Web, c.Args.Text("group"));
                group.Users.Remove(Csom.EnsureUser(ctx, c.Args.Text("who")));
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Removed from the {c.Args.Text("group").ToLowerInvariant()}.");
            },
        };

        yield return new Command
        {
            Id = "web.groups",
            Title = "List owners, members and visitors",
            Group = "Groups",
            Description = "Who is in each site's own SharePoint groups.",
            PowerShell = "Get-PnPGroupMember",
            Targets = Webs,
            ReadOnly = true,
            Glyph = "\uE716",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var groups = new[] { ("Owners", ctx.Web.AssociatedOwnerGroup), ("Members", ctx.Web.AssociatedMemberGroup), ("Visitors", ctx.Web.AssociatedVisitorGroup) };
                foreach (var (_, g) in groups)
                {
                    ctx.Load(g, x => x.Title, x => x.Users.Include(u => u.Title, u => u.Email, u => u.LoginName));
                }
                await ctx.RunAsync().ConfigureAwait(false);
                var text = new StringBuilder();
                foreach (var (name, g) in groups)
                {
                    text.AppendLine($"{name} ({g.Title}):");
                    foreach (var u in g.Users)
                    {
                        text.AppendLine($"  {u.Title}  ·  {(string.IsNullOrEmpty(u.Email) ? u.LoginName : u.Email)}");
                    }
                    if (g.Users.Count == 0)
                    {
                        text.AppendLine("  nobody");
                    }
                }
                return CommandResult.Done("The site's groups.", Join(text));
            },
        };

        foreach (var command in FeatureCommands("web", Webs, "Web", (ctx, _) => ctx.Web.Features))
        {
            yield return command;
        }

        yield return new Command
        {
            Id = "web.delete",
            Title = "Delete subsite",
            Group = "Delete",
            Description = "Deletes each subsite and everything in it (to the site collection's recycle bin). Subsites that have subsites of their own have to go first.",
            PowerShell = "Remove-PnPWeb",
            Targets = SubsitesOnly,
            Destructive = true,
            After = After.Remove,
            Glyph = "\uE74D",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                ctx.Web.DeleteObject();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Deleted.");
            },
        };
    }

    private static async Task<CommandResult> NewListAsync(CommandContext c, SpNode n, int template)
    {
        using var ctx = await c.WebAsync(n).ConfigureAwait(false);
        var list = ctx.Web.Lists.Add(new ListCreationInformation
        {
            Title = c.Args.Text("title"),
            Description = c.Args["description"],
            TemplateType = template,
        });
        list.OnQuickLaunch = c.Args.Toggle("nav");
        list.Update();
        await ctx.RunAsync().ConfigureAwait(false);
        return CommandResult.Done($"Created {c.Args.Text("title")}.");
    }
}
