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
            Description = "Whether the site's content shows up in search.",
            PowerShell = "Set-PnPWeb -NoCrawl",
            Targets = Webs,
            Glyph = "\uE721",
            Parameters = [Toggle("show", "Show in search", true)],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                ctx.Web.NoCrawl = !c.Args.Toggle("show");
                ctx.Web.Update();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done(c.Args.Toggle("show") ? "Shown in search." : "Left out of search.");
            },
        };

        yield return new Command
        {
            Id = "web.reindex",
            Title = "Reindex",
            Group = "Search",
            Description = "Asks search to crawl the whole site again at its next crawl, e.g. after changing its search settings. Needs custom scripts allowed.",
            PowerShell = "Request-PnPReindexWeb",
            Targets = Webs,
            Glyph = "\uE72C",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var props = ctx.Web.AllProperties;
                ctx.Load(props);
                await ctx.RunAsync().ConfigureAwait(false);
                var version = props.FieldValues.TryGetValue("vti_searchversion", out var v) && int.TryParse(v?.ToString(), out var i) ? i : 0;
                props["vti_searchversion"] = version + 1;
                ctx.Web.Update();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Search will crawl it again.");
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
            Description = "Stores a value in each site's property bag. Needs custom scripts allowed.",
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
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                ctx.Web.AllProperties[c.Args.Text("key")] = c.Args["value"] ?? "";
                ctx.Web.Update();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Set {c.Args.Text("key")}.");
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
