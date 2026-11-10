using System.Security;

using Microsoft.SharePoint.Client;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>Lists and libraries (Set-PnPList and friends).</summary>
public static partial class CommandCatalog
{
    private static readonly IReadOnlyList<Choice> ColumnTypes =
    [
        new("Text", "Single line of text"),
        new("Note", "Multiple lines of text"),
        new("Number", "Number"),
        new("DateTime", "Date and time"),
        new("Boolean", "Yes/no"),
        new("Choice", "Choice"),
        new("User", "Person"),
        new("URL", "Hyperlink"),
    ];

    private static IEnumerable<Command> ListCommands()
    {
        yield return new Command
        {
            Id = "list.title",
            Title = "Rename",
            Group = "Settings",
            Description = "Changes each list's title (its address stays), and its description when one is given.",
            PowerShell = "Set-PnPList -Title -Description",
            Targets = Lists,
            After = After.Update,
            Glyph = "\uE8AC",
            Parameters =
            [
                new("title", "Title", ParameterKind.Text),
                new("description", "Description", ParameterKind.LongText) { Required = false, Hint = "Empty: the description stays as it is." },
            ],
            Run = (c, n) => ChangeListAsync(c, n, l =>
            {
                l.Title = c.Args.Text("title");
                if (c.Args["description"] is { } d)
                {
                    l.Description = d;
                }
            }, "Renamed.", n with { Title = c.Args.Text("title") }),
        };

        yield return new Command
        {
            Id = "list.versioning",
            Title = "Versioning",
            Group = "Versions",
            Description = "Whether SharePoint keeps earlier versions, and how many. Minor (draft) versions are for libraries.",
            PowerShell = "Set-PnPList -EnableVersioning -MajorVersions -EnableMinorVersions",
            Targets = Lists,
            Glyph = "\uE81C",
            Parameters =
            [
                Toggle("on", "Keep versions", true),
                new("major", "Major versions to keep", ParameterKind.Number) { Default = "500", Min = 1, Max = 50000 },
                Toggle("minor", "Keep minor (draft) versions too", false, "Libraries only."),
            ],
            Run = (c, n) => ChangeListAsync(c, n, l =>
            {
                l.EnableVersioning = c.Args.Toggle("on");
                if (c.Args.Toggle("on"))
                {
                    l.MajorVersionLimit = (int)c.Args.Number("major");
                    if (n.Kind == NodeKind.Library)
                    {
                        l.EnableMinorVersions = c.Args.Toggle("minor");
                    }
                }
            }, c.Args.Toggle("on") ? $"Keeps {c.Args.Number("major"):N0} versions." : "Versioning off."),
        };

        yield return ListToggle("list.checkout", "Require check-out", "Versions", "Whether files have to be checked out before they can be edited.",
            "Set-PnPList -ForceCheckout", Libraries, "\uE8A5", "require", "Require check-out", false, (l, on) => l.ForceCheckout = on);
        yield return ListToggle("list.approval", "Content approval", "Versions", "Whether new and changed items stay drafts until someone approves them.",
            "Set-PnPList -EnableModeration", Lists, "\uE8FB", "on", "Require approval", false, (l, on) => l.EnableModeration = on);
        yield return ListToggle("list.content-types", "Content types", "Settings", "Whether the list's content types can be managed.",
            "Set-PnPList -EnableContentTypes", Lists, "\uE8EC", "on", "Allow management of content types", true, (l, on) => l.ContentTypesEnabled = on);
        yield return ListToggle("list.folders", "Folders", "Settings", "Whether the New menu offers folders.",
            "Set-PnPList -EnableFolderCreation", Lists, "\uE8F4", "on", "Allow new folders", true, (l, on) => l.EnableFolderCreation = on);
        yield return ListToggle("list.attachments", "Attachments", "Settings", "Whether items can have files attached.",
            "Set-PnPList -EnableAttachments", ListsOnly, "\uE723", "on", "Allow attachments", true, (l, on) => l.EnableAttachments = on);
        yield return ListToggle("list.navigation", "Show on the site's navigation", "Settings", "Whether the list has a link in the site's navigation.",
            "Set-PnPList -OnQuickLaunch", Lists, "\uE700", "on", "Show a link", true, (l, on) => l.OnQuickLaunch = on);
        yield return ListToggle("list.hidden", "Hide", "Settings", "Hidden lists don't show in Site contents (or in this app) but keep working.",
            "Set-PnPList -Hidden", Lists, "\uED1A", "hidden", "Hide the list", true, (l, on) => l.Hidden = on);
        yield return ListToggle("list.search", "Show in search results", "Search", "Whether the list's content shows up in search.",
            "Set-PnPList -NoCrawl", Lists, "\uE721", "show", "Show in search", true, (l, on) => l.NoCrawl = !on);

        yield return new Command
        {
            Id = "list.reindex",
            Title = "Reindex",
            Group = "Search",
            Description = "Asks search to crawl the whole list again at its next crawl.",
            PowerShell = "Request-PnPReindexList",
            Targets = Lists,
            Glyph = "\uE72C",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var folder = Csom.List(ctx, n).RootFolder;
                ctx.Load(folder.Properties);
                await ctx.RunAsync().ConfigureAwait(false);
                var version = folder.Properties.FieldValues.TryGetValue("vti_searchversion", out var v) && int.TryParse(v?.ToString(), out var i) ? i : 0;
                folder.Properties["vti_searchversion"] = version + 1;
                folder.Update();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Search will crawl it again.");
            },
        };

        yield return new Command
        {
            Id = "list.add-column",
            Title = "Add a column",
            Group = "Settings",
            Description = "Adds a column to each list or library.",
            PowerShell = "Add-PnPField",
            Targets = Lists,
            Glyph = "\uE710",
            Parameters =
            [
                new("name", "Name", ParameterKind.Text),
                new("type", "Type", ParameterKind.Choice) { Choices = ColumnTypes, Default = "Text" },
                new("choices", "Choices", ParameterKind.LongText) { Required = false, Hint = "For a choice column: one per line." },
                Toggle("view", "Add it to the default view", true),
                Toggle("required", "Require a value", false),
            ],
            Run = async (c, n) =>
            {
                var name = c.Args.Text("name");
                var type = c.Args.Text("type");
                var internalName = new string(name.Where(char.IsLetterOrDigit).ToArray());
                if (internalName.Length == 0)
                {
                    internalName = "Column";
                }
                var choices = type == "Choice"
                    ? "<CHOICES>" + string.Concat((c.Args["choices"] ?? "").Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        .Select(x => $"<CHOICE>{SecurityElement.Escape(x)}</CHOICE>")) + "</CHOICES>"
                    : "";
                var xml = $"<Field Type='{type}' DisplayName='{SecurityElement.Escape(name)}' Name='{internalName}' Required='{(c.Args.Toggle("required") ? "TRUE" : "FALSE")}'" +
                          (type == "User" ? " UserSelectionMode='PeopleOnly'" : "") + $">{choices}</Field>";
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                Csom.List(ctx, n).Fields.AddFieldAsXml(xml, c.Args.Toggle("view"), AddFieldOptions.AddFieldInternalNameHint | AddFieldOptions.AddFieldToDefaultView);
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Added the column {name}.");
            },
        };

        yield return new Command
        {
            Id = "list.delete",
            Title = "Delete list",
            Group = "Delete",
            Description = "Moves each list or library, with everything in it, to the site's recycle bin.",
            PowerShell = "Remove-PnPList -Recycle",
            Targets = Lists,
            Destructive = true,
            After = After.Remove,
            Glyph = "\uE74D",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                Csom.List(ctx, n).Recycle();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Deleted (in the recycle bin).");
            },
        };

        yield return new Command
        {
            Id = "list.delete-permanently",
            Title = "Delete list for good",
            Group = "Delete",
            Description = "Deletes each list or library and everything in it, without the recycle bin.",
            PowerShell = "Remove-PnPList",
            Targets = Lists,
            Destructive = true,
            After = After.Remove,
            Glyph = "\uE74D",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                Csom.List(ctx, n).DeleteObject();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Deleted for good.");
            },
        };
    }

    private static Command ListToggle(string id, string title, string group, string description, string powerShell, NodeKind[] targets,
        string glyph, string name, string label, bool on, Action<List, bool> apply) => new()
    {
        Id = id,
        Title = title,
        Group = group,
        Description = description,
        PowerShell = powerShell,
        Targets = targets,
        Glyph = glyph,
        Parameters = [Toggle(name, label, on)],
        Run = (c, n) => ChangeListAsync(c, n, l => apply(l, c.Args.Toggle(name)), $"{label}: {Csom.Bool(c.Args.Toggle(name))}."),
    };

    private static async Task<CommandResult> ChangeListAsync(CommandContext c, SpNode n, Action<List> change, string done, SpNode? updated = null)
    {
        using var ctx = await c.WebAsync(n).ConfigureAwait(false);
        var list = Csom.List(ctx, n);
        change(list);
        list.Update();
        await ctx.RunAsync().ConfigureAwait(false);
        return CommandResult.Done(done) with { Updated = updated };
    }
}
