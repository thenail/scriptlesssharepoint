using System.Security;

using Microsoft.SharePoint.Client;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>
/// Site columns and content types (Add-PnPField, Add-PnPContentType and
/// friends), and putting them on lists. A site's columns and content types are
/// shared by its lists and its subsites; changes can be pushed to the lists
/// that use them.
/// </summary>
public static partial class CommandCatalog
{
    private const string ColumnsGroup = "Columns and content types";

    /// <summary>The content types a new one is usually based on, by ID.</summary>
    private static readonly IReadOnlyList<Choice> ParentContentTypes =
    [
        new("0x0101", "Document"),
        new("0x01", "Item"),
        new("0x0120", "Folder"),
        new(SitePageContentType, "Site Page"),
        new("name", "Another, by name"),
    ];

    private static IEnumerable<Command> ColumnCommands()
    {
        // ---- site columns ----

        yield return new Command
        {
            Id = "web.columns",
            Title = "List site columns",
            Group = ColumnsGroup,
            Description = "The columns each site defines for its lists to use: name shown, internal name, type and group.",
            PowerShell = "Get-PnPField",
            Targets = Webs,
            ReadOnly = true,
            Glyph = "",
            Parameters = [new("group", "Group", ParameterKind.Text) { Required = false, Placeholder = "Custom Columns", Hint = "Empty: every group but SharePoint's hidden ones." }],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var fields = ctx.Web.Fields;
                ctx.Load(fields, fs => fs.Include(f => f.Title, f => f.InternalName, f => f.TypeDisplayName, f => f.Group, f => f.Hidden));
                await ctx.RunAsync().ConfigureAwait(false);
                var group = c.Args["group"]?.Trim();
                var shown = fields.ToList()
                    .Where(f => !f.Hidden && f.Group != "_Hidden" && (group is null || string.Equals(f.Group, group, StringComparison.CurrentCultureIgnoreCase)))
                    .OrderBy(f => f.Group, StringComparer.CurrentCultureIgnoreCase)
                    .ThenBy(f => f.Title, StringComparer.CurrentCultureIgnoreCase)
                    .ToList();
                return CommandResult.Done($"{Plural(shown.Count, "column", "columns")}.",
                    string.Join(Environment.NewLine, shown.Select(f => $"{f.Group}  ·  {f.Title}  ·  {f.InternalName}  ·  {f.TypeDisplayName}")));
            },
        };

        yield return new Command
        {
            Id = "web.add-column",
            Title = "New site column",
            Group = ColumnsGroup,
            Description = "Creates a column in each site that its lists and content types can use.",
            PowerShell = "Add-PnPField",
            Targets = Webs,
            Glyph = "",
            Parameters =
            [
                new("name", "Name", ParameterKind.Text),
                new("type", "Type", ParameterKind.Choice) { Choices = ColumnTypes, Default = "Text" },
                new("choices", "Choices", ParameterKind.LongText) { Hint = "One per line.", ShowIf = new ShowIf("type", "Choice") },
                new("group", "Group", ParameterKind.Text) { Default = "Custom Columns" },
                Toggle("required", "Require a value", false),
            ],
            Run = async (c, n) =>
            {
                var name = c.Args.Text("name").Trim();
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                ctx.Web.Fields.AddFieldAsXml(FieldXml(name, c.Args.Text("type"), c.Args["choices"], c.Args.Toggle("required"), c.Args.Text("group").Trim()),
                    false, AddFieldOptions.AddFieldInternalNameHint);
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Created {name}.");
            },
        };

        yield return new Command
        {
            Id = "web.change-column",
            Title = "Change a site column",
            Group = ColumnsGroup,
            Description = "Renames a site column, makes it required or not, or moves it to another group, and passes that on to the lists that use it.",
            PowerShell = "Set-PnPField -UpdateExistingLists",
            Targets = Webs,
            Glyph = "",
            Parameters =
            [
                new("name", "Column", ParameterKind.Text) { Hint = "Its name as shown, or its internal name." },
                new("title", "New name", ParameterKind.Text) { Required = false, Hint = "Empty keeps its name." },
                new("required", "Required", ParameterKind.Choice) { Default = "keep", Choices = KeepYesNo },
                new("group", "New group", ParameterKind.Text) { Required = false, Hint = "Empty keeps its group." },
                Toggle("push", "Change it in the lists that use it", true),
            ],
            Run = async (c, n) =>
            {
                if (c.Args["title"] is null && c.Args.Text("required") == "keep" && c.Args["group"] is null)
                {
                    return CommandResult.Skipped("Nothing to change.");
                }
                var name = c.Args.Text("name").Trim();
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var field = ctx.Web.Fields.GetByInternalNameOrTitle(name);
                if (await LoadOrMissingAsync(ctx, field, f => f.Title) is false)
                {
                    return CommandResult.Skipped($"It has no site column {name}.");
                }
                if (c.Args["title"] is { } title)
                {
                    field.Title = title.Trim();
                }
                if (c.Args.Text("required") != "keep")
                {
                    field.Required = c.Args.Text("required") == "yes";
                }
                if (c.Args["group"] is { } group)
                {
                    field.Group = group.Trim();
                }
                field.UpdateAndPushChanges(c.Args.Toggle("push"));
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Changed {field.Title}.");
            },
        };

        yield return new Command
        {
            Id = "web.delete-column",
            Title = "Delete a site column",
            Group = ColumnsGroup,
            Description = "Deletes a site column. Lists that use it keep their own copy of it, and its values.",
            PowerShell = "Remove-PnPField",
            Targets = Webs,
            Destructive = true,
            Glyph = "",
            Parameters = [new("name", "Column", ParameterKind.Text) { Hint = "Its name as shown, or its internal name." }],
            Run = async (c, n) =>
            {
                var name = c.Args.Text("name").Trim();
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var field = ctx.Web.Fields.GetByInternalNameOrTitle(name);
                if (await LoadOrMissingAsync(ctx, field, f => f.Title, f => f.CanBeDeleted) is false)
                {
                    return CommandResult.Skipped($"It has no site column {name}.");
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

        // ---- content types ----

        yield return new Command
        {
            Id = "web.content-types",
            Title = "List site content types",
            Group = ColumnsGroup,
            Description = "The content types each site defines: name, group and ID.",
            PowerShell = "Get-PnPContentType",
            Targets = Webs,
            ReadOnly = true,
            Glyph = "",
            Parameters = [new("group", "Group", ParameterKind.Text) { Required = false, Placeholder = "Custom Content Types", Hint = "Empty: every group but SharePoint's hidden ones." }],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var types = ctx.Web.ContentTypes;
                ctx.Load(types, ts => ts.Include(t => t.Name, t => t.Group, t => t.StringId, t => t.Hidden));
                await ctx.RunAsync().ConfigureAwait(false);
                var group = c.Args["group"]?.Trim();
                var shown = types.ToList()
                    .Where(t => !t.Hidden && t.Group != "_Hidden" && (group is null || string.Equals(t.Group, group, StringComparison.CurrentCultureIgnoreCase)))
                    .OrderBy(t => t.Group, StringComparer.CurrentCultureIgnoreCase)
                    .ThenBy(t => t.Name, StringComparer.CurrentCultureIgnoreCase)
                    .ToList();
                return CommandResult.Done($"{Plural(shown.Count, "content type", "content types")}.",
                    string.Join(Environment.NewLine, shown.Select(t => $"{t.Group}  ·  {t.Name}  ·  {t.StringId}")));
            },
        };

        yield return new Command
        {
            Id = "web.new-content-type",
            Title = "New content type",
            Group = ColumnsGroup,
            Description = "Creates a content type in each site, based on another: Document for files, Item for list items.",
            PowerShell = "Add-PnPContentType",
            Targets = Webs,
            Glyph = "",
            Parameters =
            [
                new("name", "Name", ParameterKind.Text),
                new("parent", "Based on", ParameterKind.Choice) { Choices = ParentContentTypes, Default = "0x0101" },
                new("parentName", "Based on the content type", ParameterKind.Text) { ShowIf = new ShowIf("parent", "name"), Placeholder = "Contract" },
                new("group", "Group", ParameterKind.Text) { Default = "Custom Content Types" },
                new("description", "Description", ParameterKind.LongText) { Required = false },
            ],
            Run = async (c, n) =>
            {
                var name = c.Args.Text("name").Trim();
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var parent = c.Args.Text("parent");
                if (parent == "name")
                {
                    if (await ContentTypeAsync(ctx, ctx.Web.AvailableContentTypes, c.Args.Text("parentName")) is not { } found)
                    {
                        return CommandResult.Failed($"The site has no content type called {c.Args.Text("parentName").Trim()}.");
                    }
                    parent = found.StringId;
                }
                ctx.Web.ContentTypes.Add(new ContentTypeCreationInformation
                {
                    Name = name,
                    Id = ChildContentTypeId(parent),
                    Group = c.Args.Text("group").Trim(),
                    Description = c.Args["description"] ?? "",
                });
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Created {name}.");
            },
        };

        yield return ContentTypeColumn(add: true);
        yield return ContentTypeColumn(add: false);

        yield return new Command
        {
            Id = "web.delete-content-type",
            Title = "Delete a site content type",
            Group = ColumnsGroup,
            Description = "Deletes a content type from each site. SharePoint refuses while lists or other content types use it.",
            PowerShell = "Remove-PnPContentType",
            Targets = Webs,
            Destructive = true,
            Glyph = "",
            Parameters = [new("name", "Content type", ParameterKind.Text)],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                if (await ContentTypeAsync(ctx, ctx.Web.ContentTypes, c.Args.Text("name")) is not { } type)
                {
                    return CommandResult.Skipped($"It has no content type called {c.Args.Text("name").Trim()}.");
                }
                type.DeleteObject();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Deleted {type.Name}.");
            },
        };

        // ---- on lists ----

        yield return new Command
        {
            Id = "list.add-site-column",
            Title = "Add a site column",
            Group = ColumnsGroup,
            Description = "Adds one of the site's columns to each list or library.",
            PowerShell = "Add-PnPField -List -Field",
            Targets = Lists,
            Glyph = "",
            Parameters =
            [
                new("name", "Site column", ParameterKind.Text) { Hint = "Its name as shown, or its internal name." },
                Toggle("view", "Add it to the default view", true),
            ],
            Run = async (c, n) =>
            {
                var name = c.Args.Text("name").Trim();
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var field = ctx.Web.AvailableFields.GetByInternalNameOrTitle(name);
                var list = Csom.List(ctx, n);
                if (await LoadOrMissingAsync(ctx, field, f => f.Title, f => f.InternalName) is false)
                {
                    return CommandResult.Failed($"The site has no column {name}.");
                }
                var mine = list.Fields;
                ctx.Load(mine, fs => fs.Include(f => f.InternalName));
                await ctx.RunAsync().ConfigureAwait(false);
                if (mine.ToList().Any(f => f.InternalName == field.InternalName))
                {
                    return CommandResult.Skipped($"It has {field.Title} already.");
                }
                list.Fields.Add(field);
                if (c.Args.Toggle("view"))
                {
                    list.DefaultView.ViewFields.Add(field.InternalName);
                    list.DefaultView.Update();
                }
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Added {field.Title}.");
            },
        };

        yield return new Command
        {
            Id = "list.content-type-remove",
            Title = "Remove a content type",
            Group = ColumnsGroup,
            Description = "Takes a content type off each list. SharePoint refuses while items use it.",
            PowerShell = "Remove-PnPContentTypeFromList",
            Targets = Lists,
            Glyph = "",
            Parameters = [new("name", "Content type", ParameterKind.Text)],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                if (await ContentTypeAsync(ctx, Csom.List(ctx, n).ContentTypes, c.Args.Text("name")) is not { } type)
                {
                    return CommandResult.Skipped($"It has no content type called {c.Args.Text("name").Trim()}.");
                }
                type.DeleteObject();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Removed {type.Name}.");
            },
        };

        yield return new Command
        {
            Id = "list.default-content-type",
            Title = "Default content type",
            Group = ColumnsGroup,
            Description = "Which of each list's content types New makes first.",
            PowerShell = "Set-PnPDefaultContentTypeToList",
            Targets = Lists,
            Glyph = "",
            Parameters = [new("name", "Content type", ParameterKind.Text)],
            Run = async (c, n) =>
            {
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var list = Csom.List(ctx, n);
                if (await ContentTypeAsync(ctx, list.ContentTypes, c.Args.Text("name")) is not { } type)
                {
                    return CommandResult.Skipped($"It has no content type called {c.Args.Text("name").Trim()}.");
                }
                var folder = list.RootFolder;
                ctx.Load(folder, f => f.ContentTypeOrder);
                ctx.Load(type, t => t.Id);
                await ctx.RunAsync().ConfigureAwait(false);
                var order = folder.ContentTypeOrder.Where(id => id.StringValue != type.StringId).ToList();
                order.Insert(0, type.Id);
                folder.UniqueContentTypeOrder = order;
                folder.Update();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"New makes {type.Name} first.");
            },
        };
    }

    private static Command ContentTypeColumn(bool add) => new()
    {
        Id = add ? "web.content-type-add-column" : "web.content-type-remove-column",
        Title = add ? "Add a column to a content type" : "Take a column off a content type",
        Group = ColumnsGroup,
        Description = add
            ? "Adds one of the site's columns to one of its content types, and to the lists that use the content type."
            : "Takes a column off one of the site's content types, and off the lists' copies of it (their values stay in the lists).",
        PowerShell = add ? "Add-PnPFieldToContentType" : "Remove-PnPFieldFromContentType",
        Targets = Webs,
        Glyph = add ? "" : "",
        Parameters =
        [
            new("type", "Content type", ParameterKind.Text),
            new("column", "Site column", ParameterKind.Text) { Hint = "Its name as shown, or its internal name." },
            .. add ? new[] { Toggle("required", "Require a value", false) } : [],
            Toggle("push", "Change the lists that use it too", true),
        ],
        Run = async (c, n) =>
        {
            using var ctx = await c.WebAsync(n).ConfigureAwait(false);
            if (await ContentTypeAsync(ctx, ctx.Web.ContentTypes, c.Args.Text("type")) is not { } type)
            {
                return CommandResult.Skipped($"It has no content type called {c.Args.Text("type").Trim()}.");
            }
            var name = c.Args.Text("column").Trim();
            var field = ctx.Web.AvailableFields.GetByInternalNameOrTitle(name);
            if (await LoadOrMissingAsync(ctx, field, f => f.Id, f => f.Title) is false)
            {
                return CommandResult.Failed($"The site has no column {name}.");
            }
            var links = type.FieldLinks;
            ctx.Load(links, ls => ls.Include(l => l.Id));
            await ctx.RunAsync().ConfigureAwait(false);
            var has = links.ToList().Any(l => l.Id == field.Id);
            if (add == has)
            {
                return CommandResult.Skipped(add ? $"{type.Name} has {field.Title} already." : $"{type.Name} doesn't have {field.Title}.");
            }
            if (add)
            {
                var link = links.Add(new FieldLinkCreationInformation { Field = field });
                link.Required = c.Args.Toggle("required");
            }
            else
            {
                links.GetById(field.Id).DeleteObject();
            }
            type.Update(c.Args.Toggle("push"));
            await ctx.RunAsync().ConfigureAwait(false);
            return CommandResult.Done(add ? $"Added {field.Title} to {type.Name}." : $"Took {field.Title} off {type.Name}.");
        },
    };

    // ---- helpers ----

    /// <summary>A column's XML for AddFieldAsXml: its type, name, choices, and group for a site column.</summary>
    internal static string FieldXml(string name, string type, string? choices, bool required, string? group = null)
    {
        var internalName = new string(name.Where(char.IsLetterOrDigit).ToArray());
        if (internalName.Length == 0 || char.IsDigit(internalName[0]))
        {
            internalName = "Column" + internalName;
        }
        var options = type == "Choice"
            ? "<CHOICES>" + string.Concat((choices ?? "").Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(x => $"<CHOICE>{SecurityElement.Escape(x)}</CHOICE>")) + "</CHOICES>"
            : "";
        return $"<Field ID='{{{Guid.NewGuid()}}}' Type='{type}' DisplayName='{SecurityElement.Escape(name.Trim())}' Name='{internalName}' StaticName='{internalName}'" +
               $" Required='{(required ? "TRUE" : "FALSE")}'" +
               (string.IsNullOrWhiteSpace(group) ? "" : $" Group='{SecurityElement.Escape(group)}'") +
               (type == "User" ? " UserSelectionMode='PeopleOnly'" : "") +
               $">{options}</Field>";
    }

    /// <summary>A new content type's ID under its parent's: the parent's, "00", and a GUID.</summary>
    internal static string ChildContentTypeId(string parent) => $"{parent}00{Guid.NewGuid():N}".ToUpperInvariant().Replace("0X", "0x", StringComparison.Ordinal);

    /// <summary>A content type in a collection by its name (any case), loaded with its name and ID; null when there's none.</summary>
    private static async Task<ContentType?> ContentTypeAsync(ClientContext ctx, ContentTypeCollection types, string name)
    {
        ctx.Load(types, ts => ts.Include(t => t.Name, t => t.StringId));
        await ctx.RunAsync().ConfigureAwait(false);
        return types.ToList().FirstOrDefault(t => string.Equals(t.Name, name.Trim(), StringComparison.CurrentCultureIgnoreCase));
    }

    /// <summary>Loads a column found by name; false when there's no such column.</summary>
    private static async Task<bool> LoadOrMissingAsync(ClientContext ctx, Field field, params System.Linq.Expressions.Expression<Func<Field, object>>[] properties)
    {
        ctx.Load(field, properties);
        try
        {
            await ctx.RunAsync().ConfigureAwait(false);
            return true;
        }
        catch (ServerException e) when (e.ServerErrorTypeName == "System.ArgumentException")
        {
            return false;
        }
    }
}
