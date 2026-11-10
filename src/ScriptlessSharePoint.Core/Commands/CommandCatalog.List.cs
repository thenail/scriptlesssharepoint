using System.Globalization;
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

    /// <summary>For a setting a command can leave as it is.</summary>
    private static readonly IReadOnlyList<Choice> KeepYesNo = [new("keep", "Leave as it is"), new("yes", "Yes"), new("no", "No")];

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
        yield return ListToggle("list.content-types", "Content types", ColumnsGroup, "Whether the list's content types can be managed.",
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
                var allowed = await c.WithScriptsAsync(n, async () =>
                {
                    using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                    var folder = Csom.List(ctx, n).RootFolder;
                    ctx.Load(folder.Properties);
                    await ctx.RunAsync().ConfigureAwait(false);
                    var version = folder.Properties.FieldValues.TryGetValue("vti_searchversion", out var v) && int.TryParse(v?.ToString(), out var i) ? i : 0;
                    folder.Properties["vti_searchversion"] = version + 1;
                    folder.Update();
                    await ctx.RunAsync().ConfigureAwait(false);
                }).ConfigureAwait(false);
                return CommandResult.Done(Csom.ScriptsNote("Search will crawl it again.", allowed));
            },
        };

        yield return new Command
        {
            Id = "list.add-column",
            Title = "Add a column",
            Group = ColumnsGroup,
            Description = "Adds a column to each list or library.",
            PowerShell = "Add-PnPField",
            Targets = Lists,
            Glyph = "\uE710",
            Parameters =
            [
                new("name", "Name", ParameterKind.Text),
                new("type", "Type", ParameterKind.Choice) { Choices = ColumnTypes, Default = "Text" },
                new("choices", "Choices", ParameterKind.LongText) { Hint = "One per line.", ShowIf = new ShowIf("type", "Choice") },
                Toggle("view", "Add it to the default view", true),
                Toggle("required", "Require a value", false),
            ],
            Run = async (c, n) =>
            {
                var name = c.Args.Text("name");
                var xml = FieldXml(name, c.Args.Text("type"), c.Args["choices"], c.Args.Toggle("required"));
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                Csom.List(ctx, n).Fields.AddFieldAsXml(xml, c.Args.Toggle("view"), AddFieldOptions.AddFieldInternalNameHint | AddFieldOptions.AddFieldToDefaultView);
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Added the column {name}.");
            },
        };

        yield return new Command
        {
            Id = "list.columns",
            Title = "List columns",
            Group = ColumnsGroup,
            Description = "Each list's columns: the name shown, the internal name (for Set a column's value) and the type.",
            PowerShell = "Get-PnPField",
            Targets = Lists,
            ReadOnly = true,
            Glyph = "\uE8FD",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var fields = Csom.List(ctx, n).Fields;
                ctx.Load(fields, fs => fs.Include(f => f.Title, f => f.InternalName, f => f.TypeDisplayName, f => f.Hidden));
                await ctx.RunAsync().ConfigureAwait(false);
                var shown = fields.ToList().Where(f => !f.Hidden).OrderBy(f => f.Title, StringComparer.CurrentCultureIgnoreCase).ToList();
                return CommandResult.Report($"{shown.Count} columns.",
                    ReportTable.Of(["Column", "Internal name", "Type"], shown.Select(f => new[] { f.Title, f.InternalName, f.TypeDisplayName })));
            },
        };

        yield return new Command
        {
            Id = "list.delete-column",
            Title = "Delete a column",
            Group = ColumnsGroup,
            Description = "Deletes a column, and its values, from each list or library.",
            PowerShell = "Remove-PnPField",
            Targets = Lists,
            Destructive = true,
            Glyph = "\uE74D",
            Parameters = [new("name", "Column", ParameterKind.Text) { Hint = "Its name as shown, or its internal name." }],
            Run = async (c, n) =>
            {
                var name = c.Args.Text("name").Trim();
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var field = Csom.List(ctx, n).Fields.GetByInternalNameOrTitle(name);
                ctx.Load(field, f => f.CanBeDeleted, f => f.Title);
                try
                {
                    await ctx.RunAsync().ConfigureAwait(false);
                }
                catch (ServerException e) when (e.ServerErrorTypeName == "System.ArgumentException")
                {
                    return CommandResult.Skipped($"It has no column {name}.");
                }
                if (!field.CanBeDeleted)
                {
                    return CommandResult.Failed($"SharePoint doesn't let {field.Title} be deleted.");
                }
                field.DeleteObject();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Deleted {field.Title}.");
            },
        };

        yield return new Command
        {
            Id = "list.experience",
            Title = "Modern or classic",
            Group = "Settings",
            Description = "Whether each list opens in the modern or the classic experience.",
            PowerShell = "Set-PnPList -ListExperience",
            Targets = Lists,
            Glyph = "\uE7F4",
            Parameters =
            [
                new("experience", "Opens in", ParameterKind.Choice)
                {
                    Default = nameof(ListExperience.NewExperience),
                    Choices =
                    [
                        new(nameof(ListExperience.NewExperience), "Modern"),
                        new(nameof(ListExperience.ClassicExperience), "Classic"),
                        new(nameof(ListExperience.Auto), "As the site decides"),
                    ],
                },
            ],
            Run = (c, n) => ChangeListAsync(c, n, l => l.ListExperienceOptions = c.Args.Enum<ListExperience>("experience"), "Changed."),
        };

        yield return new Command
        {
            Id = "list.open-in",
            Title = "Open documents in",
            Group = "Settings",
            Description = "Whether documents in each library open in the browser or in the desktop app.",
            PowerShell = "Set-PnPList -OpenDocumentsMode",
            Targets = Libraries,
            Glyph = "\uE8A7",
            Parameters = [new("mode", "Open in", ParameterKind.Choice) { Default = "browser", Choices = [new("browser", "The browser"), new("client", "The desktop app")] }],
            Run = (c, n) => ChangeListAsync(c, n, l => l.DefaultItemOpenInBrowser = c.Args.Text("mode") == "browser",
                c.Args.Text("mode") == "browser" ? "Opens in the browser." : "Opens in the desktop app."),
        };

        yield return new Command
        {
            Id = "list.drafts",
            Title = "Who sees drafts",
            Group = "Versions",
            Description = "Who can see draft (minor or unapproved) versions in each list or library.",
            PowerShell = "Set-PnPList -DraftVersionVisibility",
            Targets = Lists,
            Glyph = "\uE7B3",
            Parameters =
            [
                new("who", "Drafts are seen by", ParameterKind.Choice)
                {
                    Default = nameof(DraftVisibilityType.Author),
                    Choices =
                    [
                        new(nameof(DraftVisibilityType.Reader), "Anyone who can read the list"),
                        new(nameof(DraftVisibilityType.Author), "Only people who can edit"),
                        new(nameof(DraftVisibilityType.Approver), "Only people who can approve, and the author"),
                    ],
                },
            ],
            Run = (c, n) => ChangeListAsync(c, n, l => l.DraftVersionVisibility = c.Args.Enum<DraftVisibilityType>("who"), "Changed who sees drafts."),
        };

        yield return new Command
        {
            Id = "list.item-access",
            Title = "Item-level access",
            Group = "Permissions",
            Description = "Whether people can read and edit everyone's items in each list, or only their own (the list's advanced settings).",
            PowerShell = "Set-PnPList -ReadSecurity -WriteSecurity",
            Targets = ListsOnly,
            Glyph = "\uE8D7",
            Parameters =
            [
                new("read", "Read", ParameterKind.Choice) { Default = "1", Choices = [new("1", "All items"), new("2", "Only their own")] },
                new("write", "Create and edit", ParameterKind.Choice) { Default = "1", Choices = [new("1", "All items"), new("2", "Only their own"), new("4", "None")] },
            ],
            Run = (c, n) => ChangeListAsync(c, n, l =>
            {
                l.ReadSecurity = int.Parse(c.Args.Text("read"), CultureInfo.InvariantCulture);
                l.WriteSecurity = int.Parse(c.Args.Text("write"), CultureInfo.InvariantCulture);
            }, "Changed item-level access."),
        };

        yield return new Command
        {
            Id = "list.label",
            Title = "Default sensitivity label",
            Group = "Settings",
            Description = "The sensitivity label files in each library get unless they have one. Empty takes it off.",
            PowerShell = "Set-PnPList -DefaultSensitivityLabelForLibrary",
            Targets = Libraries,
            Glyph = "\uEA18",
            Parameters =
            [
                new("label", "Label ID", ParameterKind.Text)
                {
                    Required = false,
                    Placeholder = "27451a5b-5823-4853-bcd4-2204d03ab477",
                    Hint = "The label's ID (a GUID), from Purview or Get-Label. Empty takes it off.",
                },
            ],
            Run = (c, n) => LabelId(c.Args["label"]) is not { } label
                ? Task.FromResult(CommandResult.Failed("The label ID isn't a GUID."))
                : ChangeListAsync(c, n, l => l.DefaultSensitivityLabelForLibrary = label, label.Length == 0 ? "Label taken off." : "Label set."),
        };

        yield return new Command
        {
            Id = "list.column-settings",
            Title = "Change a column",
            Group = ColumnsGroup,
            Description = "Renames a column of each list, makes it required or not, and indexes it or not (an index lets lists over 5,000 items filter and sort by it).",
            PowerShell = "Set-PnPField",
            Targets = Lists,
            Glyph = "\uE70F",
            Parameters =
            [
                new("name", "Column", ParameterKind.Text) { Hint = "Its name as shown, or its internal name." },
                new("title", "New name", ParameterKind.Text) { Required = false, Hint = "Empty keeps its name." },
                new("required", "Required", ParameterKind.Choice) { Default = "keep", Choices = KeepYesNo },
                new("indexed", "Indexed", ParameterKind.Choice) { Default = "keep", Choices = KeepYesNo },
            ],
            Run = async (c, n) =>
            {
                var name = c.Args.Text("name").Trim();
                if (c.Args["title"] is null && c.Args.Text("required") == "keep" && c.Args.Text("indexed") == "keep")
                {
                    return CommandResult.Skipped("Nothing to change.");
                }
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var field = Csom.List(ctx, n).Fields.GetByInternalNameOrTitle(name);
                ctx.Load(field, f => f.Title);
                try
                {
                    await ctx.RunAsync().ConfigureAwait(false);
                }
                catch (ServerException e) when (e.ServerErrorTypeName == "System.ArgumentException")
                {
                    return CommandResult.Skipped($"It has no column {name}.");
                }
                if (c.Args["title"] is { } title)
                {
                    field.Title = title.Trim();
                }
                if (c.Args.Text("required") != "keep")
                {
                    field.Required = c.Args.Text("required") == "yes";
                }
                if (c.Args.Text("indexed") != "keep")
                {
                    field.Indexed = c.Args.Text("indexed") == "yes";
                }
                field.Update();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Changed {field.Title}.");
            },
        };

        yield return new Command
        {
            Id = "list.views",
            Title = "List views",
            Group = "Details",
            Description = "Each list's views, the default one marked.",
            PowerShell = "Get-PnPView",
            Targets = Lists,
            ReadOnly = true,
            Glyph = "\uE8A9",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var views = Csom.List(ctx, n).Views;
                ctx.Load(views, vs => vs.Include(v => v.Title, v => v.DefaultView, v => v.RowLimit, v => v.PersonalView, v => v.Hidden));
                await ctx.RunAsync().ConfigureAwait(false);
                var shown = views.ToList().Where(v => !v.Hidden && !string.IsNullOrEmpty(v.Title)).ToList();
                return CommandResult.Report($"{Plural(shown.Count, "view", "views")}.", ReportTable.Of(["View", "Default", "Personal", "Page size"], shown.Select(v => new[]
                {
                    v.Title, v.DefaultView ? "default" : "", v.PersonalView ? "personal" : "", $"{v.RowLimit} a page",
                })));
            },
        };

        yield return new Command
        {
            Id = "list.add-view",
            Title = "Add a view",
            Group = "Settings",
            Description = "Adds a view to each list, with the columns given (or the default view's).",
            PowerShell = "Add-PnPView",
            Targets = Lists,
            Glyph = "\uE710",
            Parameters =
            [
                new("title", "Name", ParameterKind.Text),
                new("fields", "Columns", ParameterKind.Text) { Required = false, Placeholder = "LinkTitle, Modified, Editor", Hint = "Internal names, separated by commas (List columns shows them). Empty: the default view's." },
                new("rows", "Items a page", ParameterKind.Number) { Default = "30", Min = 1, Max = 5000 },
                Toggle("default", "Make it the default view", false),
            ],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var list = Csom.List(ctx, n);
                string[] fields;
                if (c.Args["fields"] is { } typed)
                {
                    fields = typed.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                }
                else
                {
                    var current = list.DefaultView.ViewFields;
                    ctx.Load(current);
                    await ctx.RunAsync().ConfigureAwait(false);
                    fields = current.ToArray();
                }
                list.Views.Add(new ViewCreationInformation
                {
                    Title = c.Args.Text("title").Trim(),
                    ViewFields = fields,
                    RowLimit = (uint)c.Args.Number("rows"),
                    Paged = true,
                    SetAsDefaultView = c.Args.Toggle("default"),
                    ViewTypeKind = ViewType.Html,
                });
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Added {c.Args.Text("title").Trim()}.");
            },
        };

        yield return new Command
        {
            Id = "list.content-type-list",
            Title = "List content types",
            Group = ColumnsGroup,
            Description = "The content types each list uses.",
            PowerShell = "Get-PnPContentType -List",
            Targets = Lists,
            ReadOnly = true,
            Glyph = "\uE8EC",
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var types = Csom.List(ctx, n).ContentTypes;
                ctx.Load(types, ts => ts.Include(t => t.Name, t => t.Hidden, t => t.StringId));
                await ctx.RunAsync().ConfigureAwait(false);
                var shown = types.ToList().Where(t => !t.Hidden).ToList();
                return CommandResult.Report($"{Plural(shown.Count, "content type", "content types")}.",
                    ReportTable.Of(["Content type", "ID"], shown.Select(t => new[] { t.Name, t.StringId })));
            },
        };

        yield return new Command
        {
            Id = "list.content-type-add",
            Title = "Add a content type",
            Group = ColumnsGroup,
            Description = "Adds one of the site's content types to each list, turning on content types where they're off.",
            PowerShell = "Add-PnPContentTypeToList",
            Targets = Lists,
            Glyph = "\uE710",
            Parameters = [new("name", "Content type", ParameterKind.Text) { Placeholder = "Contract", Hint = "Its name, from the site's (or the hub's) content types." }],
            Run = async (c, n) =>
            {
                var name = c.Args.Text("name").Trim();
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var list = Csom.List(ctx, n);
                var available = ctx.Web.AvailableContentTypes;
                ctx.Load(available, ts => ts.Include(t => t.Name));
                ctx.Load(list.ContentTypes, ts => ts.Include(t => t.Name));
                await ctx.RunAsync().ConfigureAwait(false);
                if (list.ContentTypes.ToList().Any(t => string.Equals(t.Name, name, StringComparison.CurrentCultureIgnoreCase)))
                {
                    return CommandResult.Skipped($"It has {name} already.");
                }
                if (available.ToList().FirstOrDefault(t => string.Equals(t.Name, name, StringComparison.CurrentCultureIgnoreCase)) is not { } type)
                {
                    return CommandResult.Failed($"The site has no content type called {name}.");
                }
                list.ContentTypesEnabled = true;
                list.Update();
                list.ContentTypes.AddExistingContentType(type);
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Added {type.Name}.");
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
