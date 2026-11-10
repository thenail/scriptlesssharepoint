using System.Security;

using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>
/// The term store (managed metadata): its groups, term sets and terms, the
/// columns that use them, and setting terms on items (the PnP *-PnPTerm*
/// cmdlets). The store is the tenant's; changing it needs a term store
/// administrator (or a group's manager or contributor). A term is named by its
/// path from the term set: "Europe/Denmark".
/// </summary>
public static partial class CommandCatalog
{
    private const string TermsGroup = "Term store";

    private static CommandParameter TermGroupName(string label = "Term group") =>
        new("group", label, ParameterKind.Text) { Placeholder = "Contoso" };

    private static CommandParameter TermSetName() =>
        new("set", "Term set", ParameterKind.Text) { Placeholder = "Countries" };

    private static CommandParameter TermPath(string label = "Term") =>
        new("term", label, ParameterKind.Text) { Placeholder = "Europe/Denmark", Hint = "Its path in the term set: parent terms first, separated by /." };

    private static IEnumerable<Command> TermCommands()
    {
        yield return new Command
        {
            Id = "tenant.term-store",
            Title = "Show the term store",
            Group = TermsGroup,
            Description = "The term store's groups and their term sets, with IDs. Site collections' own groups are marked.",
            PowerShell = "Get-PnPTermGroup, Get-PnPTermSet",
            Targets = TenantOnly,
            ReadOnly = true,
            Glyph = "",
            Parameters = [new("group", "Term group", ParameterKind.Text) { Required = false, Hint = "Empty: every group." }],
            Run = async (c, _) =>
            {
                var (ctx, store) = await TermStoreAsync(c, null).ConfigureAwait(false);
                using (ctx)
                {
                    var groups = store.Groups;
                    ctx.Load(groups, gs => gs.Include(g => g.Name, g => g.Id, g => g.IsSystemGroup, g => g.IsSiteCollectionGroup,
                        g => g.TermSets.Include(s => s.Name, s => s.Id, s => s.IsOpenForTermCreation)));
                    await ctx.RunAsync().ConfigureAwait(false);
                    var wanted = c.Args["group"];
                    var shown = groups.ToList()
                        .Where(g => !g.IsSystemGroup && (wanted is null || SameTerm(g.Name, wanted)))
                        .OrderBy(g => g.IsSiteCollectionGroup)
                        .ThenBy(g => g.Name, StringComparer.CurrentCultureIgnoreCase)
                        .ToList();
                    var lines = shown.SelectMany(g => new[] { $"{g.Name}{(g.IsSiteCollectionGroup ? "  ·  a site collection's" : "")}  ·  {g.Id}" }
                        .Concat(g.TermSets.ToList().OrderBy(s => s.Name, StringComparer.CurrentCultureIgnoreCase)
                            .Select(s => $"    {s.Name}{(s.IsOpenForTermCreation ? "  ·  open" : "")}  ·  {s.Id}")));
                    return CommandResult.Done($"{Plural(shown.Count, "group", "groups")}, {Plural(shown.Sum(g => g.TermSets.Count), "term set", "term sets")}.",
                        string.Join(Environment.NewLine, lines));
                }
            },
        };

        yield return new Command
        {
            Id = "tenant.terms",
            Title = "Show a term set's terms",
            Group = TermsGroup,
            Description = "Every term in a term set, by its path, with IDs. Deprecated terms are marked.",
            PowerShell = "Get-PnPTerm",
            Targets = TenantOnly,
            ReadOnly = true,
            Glyph = "",
            Parameters = [TermGroupName(), TermSetName()],
            Run = (c, _) => WithTermSetAsync(c, async (ctx, store, set) =>
            {
                var terms = set.GetAllTermsIncludeDeprecated();
                ctx.Load(terms, ts => ts.Include(t => t.PathOfTerm, t => t.Id, t => t.IsDeprecated, t => t.IsAvailableForTagging));
                await ctx.RunAsync().ConfigureAwait(false);
                var lines = terms.ToList()
                    .OrderBy(t => t.PathOfTerm, StringComparer.CurrentCultureIgnoreCase)
                    .Select(t => $"{t.PathOfTerm.Replace(";", " / ", StringComparison.Ordinal)}{(t.IsDeprecated ? "  ·  deprecated" : "")}{(!t.IsAvailableForTagging ? "  ·  not for tagging" : "")}  ·  {t.Id}")
                    .ToList();
                return CommandResult.Done($"{Plural(lines.Count, "term", "terms")}.", string.Join(Environment.NewLine, lines));
            }),
        };

        yield return new Command
        {
            Id = "tenant.term-group-new",
            Title = "New term group",
            Group = TermsGroup,
            Description = "Creates a group in the term store, for term sets.",
            PowerShell = "New-PnPTermGroup",
            Targets = TenantOnly,
            Glyph = "",
            Parameters = [new("name", "Name", ParameterKind.Text), new("description", "Description", ParameterKind.LongText) { Required = false }],
            Run = async (c, _) =>
            {
                var name = c.Args.Text("name").Trim();
                var (ctx, store) = await TermStoreAsync(c, null).ConfigureAwait(false);
                using (ctx)
                {
                    if (await TermGroupAsync(ctx, store, name).ConfigureAwait(false) is not null)
                    {
                        return CommandResult.Skipped($"The term store has a group called {name} already.");
                    }
                    var group = store.CreateGroup(name, Guid.NewGuid());
                    if (c.Args["description"] is { } description)
                    {
                        group.Description = description;
                    }
                    store.CommitAll();
                    await ctx.RunAsync().ConfigureAwait(false);
                    return CommandResult.Done($"Created {name}.");
                }
            },
        };

        yield return new Command
        {
            Id = "tenant.term-set-new",
            Title = "New term set",
            Group = TermsGroup,
            Description = "Creates a term set in one of the term store's groups.",
            PowerShell = "New-PnPTermSet",
            Targets = TenantOnly,
            Glyph = "",
            Parameters =
            [
                TermGroupName(),
                new("name", "Name", ParameterKind.Text),
                new("description", "Description", ParameterKind.LongText) { Required = false },
                Toggle("open", "People can add terms when they tag", false),
            ],
            Run = async (c, _) =>
            {
                var name = c.Args.Text("name").Trim();
                var (ctx, store) = await TermStoreAsync(c, null).ConfigureAwait(false);
                using (ctx)
                {
                    if (await TermGroupAsync(ctx, store, c.Args.Text("group")).ConfigureAwait(false) is not { } group)
                    {
                        return CommandResult.Failed($"The term store has no group called {c.Args.Text("group").Trim()}.");
                    }
                    if (await TermSetInAsync(ctx, group, name).ConfigureAwait(false) is not null)
                    {
                        return CommandResult.Skipped($"{group.Name} has a term set called {name} already.");
                    }
                    var set = group.CreateTermSet(name, Guid.NewGuid(), store.DefaultLanguage);
                    set.IsOpenForTermCreation = c.Args.Toggle("open");
                    if (c.Args["description"] is { } description)
                    {
                        set.Description = description;
                    }
                    store.CommitAll();
                    await ctx.RunAsync().ConfigureAwait(false);
                    return CommandResult.Done($"Created {name} in {group.Name}.");
                }
            },
        };

        yield return new Command
        {
            Id = "tenant.terms-add",
            Title = "Add terms",
            Group = TermsGroup,
            Description = "Adds terms to a term set, one per line, by their paths: \"Europe/Denmark\" adds Denmark under Europe, and Europe too if it isn't there. Terms there already are left as they are.",
            PowerShell = "New-PnPTerm, Import-PnPTermSet",
            Targets = TenantOnly,
            Glyph = "",
            Parameters =
            [
                TermGroupName(),
                TermSetName(),
                new("terms", "Terms", ParameterKind.LongText) { Placeholder = "Europe/Denmark\nEurope/Sweden\nAsia/Japan", Hint = "One per line; / separates a term from the one it's under." },
            ],
            Run = (c, _) => WithTermSetAsync(c, async (ctx, store, set) =>
            {
                var paths = c.Args.Text("terms").Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Select(TermPathOf).Where(p => p.Length > 0).ToList();
                var tree = new TermTree(ctx, set, store.DefaultLanguage);
                var (created, there) = (0, 0);
                foreach (var path in paths)
                {
                    var (_, made) = await tree.EnsureAsync(path).ConfigureAwait(false);
                    created += made;
                    there += made == 0 ? 1 : 0;
                }
                if (created == 0)
                {
                    return CommandResult.Skipped("Every term was there already.");
                }
                store.CommitAll();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Added {Plural(created, "term", "terms")}{(there > 0 ? $"; {there:N0} were there already" : "")}.");
            }),
        };

        yield return new Command
        {
            Id = "tenant.term-rename",
            Title = "Rename a term",
            Group = TermsGroup,
            Description = "Gives a term a new name. Items tagged with it show the new name as SharePoint updates them.",
            PowerShell = "Set-PnPTerm -Name",
            Targets = TenantOnly,
            Glyph = "",
            Parameters = [TermGroupName(), TermSetName(), TermPath(), new("name", "New name", ParameterKind.Text)],
            Run = (c, _) => WithTermAsync(c, (term, store) =>
            {
                term.Name = c.Args.Text("name").Trim();
                return $"Renamed to {c.Args.Text("name").Trim()}.";
            }),
        };

        yield return new Command
        {
            Id = "tenant.term-deprecate",
            Title = "Deprecate a term",
            Group = TermsGroup,
            Description = "Deprecated terms stay on what's tagged with them but can't be picked any more. Or brings one back.",
            PowerShell = "Set-PnPTerm -Deprecated",
            Targets = TenantOnly,
            Glyph = "",
            Parameters = [TermGroupName(), TermSetName(), TermPath(), Toggle("deprecate", "Deprecate it", true, "Off: it can be picked again.")],
            Run = (c, _) => WithTermAsync(c, (term, store) =>
            {
                term.Deprecate(c.Args.Toggle("deprecate"));
                return c.Args.Toggle("deprecate") ? "Deprecated." : "It can be picked again.";
            }),
        };

        yield return new Command
        {
            Id = "tenant.term-delete",
            Title = "Delete a term",
            Group = TermsGroup,
            Description = "Deletes a term and the terms under it. Items tagged with them keep the text, no longer linked to the term store.",
            PowerShell = "Remove-PnPTerm",
            Targets = TenantOnly,
            Destructive = true,
            Glyph = "",
            Parameters = [TermGroupName(), TermSetName(), TermPath()],
            Run = (c, _) => WithTermAsync(c, (term, store) =>
            {
                term.DeleteObject();
                return "Deleted.";
            }),
        };

        yield return new Command
        {
            Id = "tenant.term-set-delete",
            Title = "Delete a term set",
            Group = TermsGroup,
            Description = "Deletes a term set and every term in it. Columns using it stop offering terms.",
            PowerShell = "Remove-PnPTermSet",
            Targets = TenantOnly,
            Destructive = true,
            Glyph = "",
            Parameters = [TermGroupName(), TermSetName()],
            Run = (c, _) => WithTermSetAsync(c, async (ctx, store, set) =>
            {
                set.DeleteObject();
                store.CommitAll();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done("Deleted.");
            }),
        };

        yield return new Command
        {
            Id = "tenant.term-group-delete",
            Title = "Delete a term group",
            Group = TermsGroup,
            Description = "Deletes an empty group from the term store. SharePoint refuses while it has term sets.",
            PowerShell = "Remove-PnPTermGroup",
            Targets = TenantOnly,
            Destructive = true,
            Glyph = "",
            Parameters = [TermGroupName()],
            Run = async (c, _) =>
            {
                var (ctx, store) = await TermStoreAsync(c, null).ConfigureAwait(false);
                using (ctx)
                {
                    if (await TermGroupAsync(ctx, store, c.Args.Text("group")).ConfigureAwait(false) is not { } group)
                    {
                        return CommandResult.Skipped($"The term store has no group called {c.Args.Text("group").Trim()}.");
                    }
                    group.DeleteObject();
                    store.CommitAll();
                    await ctx.RunAsync().ConfigureAwait(false);
                    return CommandResult.Done($"Deleted {group.Name}.");
                }
            },
        };

        // ---- columns and items ----

        yield return TermColumn(list: false);
        yield return TermColumn(list: true);

        yield return new Command
        {
            Id = "content.set-term",
            Title = "Set a managed metadata value",
            Group = "Content",
            Description = "Tags each item, file or folder with terms, in one of its list's managed metadata columns.",
            PowerShell = "Set-PnPTaxonomyFieldValue",
            Targets = Content,
            Glyph = "",
            Parameters =
            [
                new("column", "Column", ParameterKind.Text) { Hint = "Its name as shown, or its internal name." },
                new("terms", "Terms", ParameterKind.Text) { Placeholder = "Europe/Denmark", Hint = "Paths in the column's term set; several (for a column that takes them) separated by commas." },
            ],
            Run = async (c, n) =>
            {
                var name = c.Args.Text("column").Trim();
                using var ctx = await c.WebAsync(n).ConfigureAwait(false);
                var field = Csom.List(ctx, n).Fields.GetByInternalNameOrTitle(name);
                if (await LoadOrMissingAsync(ctx, field, f => f.TypeAsString, f => f.InternalName) is false)
                {
                    return CommandResult.Skipped($"Its list has no column {name}.");
                }
                if (!field.TypeAsString.StartsWith("TaxonomyFieldType", StringComparison.Ordinal))
                {
                    return CommandResult.Failed($"{name} isn't a managed metadata column.");
                }
                var tax = ctx.CastTo<TaxonomyField>(field);
                ctx.Load(tax, t => t.SspId, t => t.TermSetId, t => t.AllowMultipleValues);
                await ctx.RunAsync().ConfigureAwait(false);
                var paths = c.Args.Text("terms").Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Select(TermPathOf).Where(p => p.Length > 0).ToList();
                if (paths.Count > 1 && !tax.AllowMultipleValues)
                {
                    return CommandResult.Failed($"{name} takes one term.");
                }
                var store = TaxonomySession.GetTaxonomySession(ctx).TermStores.GetById(tax.SspId);
                var tree = new TermTree(ctx, store.GetTermSet(tax.TermSetId), 1033);
                var terms = new List<Term>();
                foreach (var path in paths)
                {
                    if (await tree.FindAsync(path).ConfigureAwait(false) is not { } term)
                    {
                        return CommandResult.Failed($"The column's term set has no term {string.Join("/", path)}.");
                    }
                    terms.Add(term);
                }
                var item = Csom.Item(ctx, n);
                if (tax.AllowMultipleValues)
                {
                    var value = string.Join(";#", terms.Select(t => $"-1;#{t.Name}|{t.Id}"));
                    tax.SetFieldValueByValueCollection(item, new TaxonomyFieldValueCollection(ctx, value, tax));
                }
                else
                {
                    tax.SetFieldValueByTerm(item, terms[0], 1033);
                }
                item.Update();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Tagged with {string.Join(", ", terms.Select(t => t.Name))}.");
            },
        };
    }

    private static Command TermColumn(bool list) => new()
    {
        Id = list ? "list.add-term-column" : "web.add-term-column",
        Title = list ? "Add a managed metadata column" : "New managed metadata site column",
        Group = ColumnsGroup,
        Description = (list ? "Adds a column to each list or library" : "Creates a site column in each site") + " whose values are terms from a term set.",
        PowerShell = list ? "Add-PnPTaxonomyField -List" : "Add-PnPTaxonomyField",
        Targets = list ? Lists : Webs,
        Glyph = "",
        Parameters =
        [
            new("name", "Name", ParameterKind.Text),
            TermGroupName(),
            TermSetName(),
            Toggle("multi", "Allow several terms", false),
            Toggle("required", "Require a value", false),
            list ? Toggle("view", "Add it to the default view", true) : new("columnGroup", "Site column group", ParameterKind.Text) { Default = "Custom Columns" },
        ],
        Run = async (c, n) =>
        {
            var name = c.Args.Text("name").Trim();
            var (ctx, store) = await TermStoreAsync(c, n).ConfigureAwait(false);
            using (ctx)
            {
                if (await TermSetAsync(ctx, store, c.Args.Text("group"), c.Args.Text("set")).ConfigureAwait(false) is not { } set)
                {
                    return CommandResult.Failed($"The term store has no term set {c.Args.Text("set").Trim()} in {c.Args.Text("group").Trim()}.");
                }
                var fields = list ? Csom.List(ctx, n).Fields : ctx.Web.Fields;
                var view = list && c.Args.Toggle("view");
                var (note, field) = TaxonomyFieldXml(name, c.Args.Toggle("multi"), c.Args.Toggle("required"), list ? null : c.Args["columnGroup"]);
                fields.AddFieldAsXml(note, false, AddFieldOptions.AddFieldInternalNameHint);
                var created = fields.AddFieldAsXml(field, view, AddFieldOptions.AddFieldInternalNameHint | (view ? AddFieldOptions.AddFieldToDefaultView : 0));
                var tax = ctx.CastTo<TaxonomyField>(created);
                tax.SspId = store.Id;
                tax.TermSetId = set.Id;
                tax.TargetTemplate = "";
                tax.AnchorId = Guid.Empty;
                tax.Update();
                await ctx.RunAsync().ConfigureAwait(false);
                return CommandResult.Done($"Added {name}, for terms from {set.Name}.");
            }
        },
    };

    // ---- the term store ----

    /// <summary>The tenant's term store, seen from a site (or the tenant's root).</summary>
    private static async Task<(ClientContext Ctx, TermStore Store)> TermStoreAsync(CommandContext c, SpNode? n)
    {
        var ctx = await c.Client.ContextAsync(n?.WebUrl ?? c.Client.TenantUrl, c.Cancellation).ConfigureAwait(false);
        var store = TaxonomySession.GetTaxonomySession(ctx).GetDefaultSiteCollectionTermStore();
        ctx.Load(store, s => s.Id, s => s.Name, s => s.DefaultLanguage);
        await ctx.RunAsync().ConfigureAwait(false);
        if (store.ServerObjectIsNull == true)
        {
            ctx.Dispose();
            throw new SharePointOperationException("SharePoint has no term store for this tenant.");
        }
        return (ctx, store);
    }

    private static async Task<TermGroup?> TermGroupAsync(ClientContext ctx, TermStore store, string name)
    {
        ctx.Load(store.Groups, gs => gs.Include(g => g.Name, g => g.Id));
        await ctx.RunAsync().ConfigureAwait(false);
        return store.Groups.ToList().FirstOrDefault(g => SameTerm(g.Name, name));
    }

    private static async Task<TermSet?> TermSetInAsync(ClientContext ctx, TermGroup group, string name)
    {
        ctx.Load(group.TermSets, ss => ss.Include(s => s.Name, s => s.Id));
        await ctx.RunAsync().ConfigureAwait(false);
        return group.TermSets.ToList().FirstOrDefault(s => SameTerm(s.Name, name));
    }

    private static async Task<TermSet?> TermSetAsync(ClientContext ctx, TermStore store, string group, string set) =>
        await TermGroupAsync(ctx, store, group).ConfigureAwait(false) is { } g ? await TermSetInAsync(ctx, g, set).ConfigureAwait(false) : null;

    /// <summary>Runs with the term set the "group" and "set" inputs name.</summary>
    private static async Task<CommandResult> WithTermSetAsync(CommandContext c, Func<ClientContext, TermStore, TermSet, Task<CommandResult>> act)
    {
        var (ctx, store) = await TermStoreAsync(c, null).ConfigureAwait(false);
        using (ctx)
        {
            if (await TermSetAsync(ctx, store, c.Args.Text("group"), c.Args.Text("set")).ConfigureAwait(false) is not { } set)
            {
                return CommandResult.Failed($"The term store has no term set {c.Args.Text("set").Trim()} in {c.Args.Text("group").Trim()}.");
            }
            return await act(ctx, store, set).ConfigureAwait(false);
        }
    }

    /// <summary>Changes the term the "term" input names, then commits.</summary>
    private static Task<CommandResult> WithTermAsync(CommandContext c, Func<Term, TermStore, string> change) =>
        WithTermSetAsync(c, async (ctx, store, set) =>
        {
            var path = TermPathOf(c.Args.Text("term"));
            if (await new TermTree(ctx, set, store.DefaultLanguage).FindAsync(path).ConfigureAwait(false) is not { } term)
            {
                return CommandResult.Failed($"{set.Name} has no term {string.Join("/", path)}.");
            }
            var done = change(term, store);
            store.CommitAll();
            await ctx.RunAsync().ConfigureAwait(false);
            return CommandResult.Done(done);
        });

    /// <summary>"Europe / Denmark" (or "Europe;Denmark") as its parts.</summary>
    internal static string[] TermPathOf(string typed) =>
        typed.Split(['/', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    /// <summary>
    /// Whether two term store names are the same: SharePoint keeps &amp; and "
    /// as their full-width forms, and names aren't case-sensitive.
    /// </summary>
    internal static bool SameTerm(string a, string b) =>
        string.Equals(NormalTerm(a), NormalTerm(b), StringComparison.CurrentCultureIgnoreCase);

    private static string NormalTerm(string name) => string.Join(" ", name.Trim().Replace('&', '＆').Replace('"', '＂')
        .Split(' ', StringSplitOptions.RemoveEmptyEntries));

    /// <summary>The hidden note column a managed metadata column keeps its text in, and the column itself, as PnP makes them.</summary>
    internal static (string Note, string Field) TaxonomyFieldXml(string name, bool multi, bool required, string? group)
    {
        var internalName = new string(name.Where(char.IsLetterOrDigit).ToArray());
        if (internalName.Length == 0 || char.IsDigit(internalName[0]))
        {
            internalName = "Column" + internalName;
        }
        var noteId = Guid.NewGuid().ToString("B");
        var note = $"<Field Type='Note' DisplayName='{internalName}_0' Name='{internalName}TaxHTField0' StaticName='{internalName}TaxHTField0' ID='{noteId}'" +
                   " ShowInViewForms='FALSE' Required='FALSE' Hidden='TRUE' CanToggleHidden='TRUE' />";
        var field = $"<Field Type='{(multi ? "TaxonomyFieldTypeMulti" : "TaxonomyFieldType")}' DisplayName='{SecurityElement.Escape(name)}' Name='{internalName}'" +
                    $" StaticName='{internalName}' ID='{Guid.NewGuid():B}' ShowField='Term1033' Required='{(required ? "TRUE" : "FALSE")}'" +
                    (multi ? " Mult='TRUE'" : "") +
                    (string.IsNullOrWhiteSpace(group) ? "" : $" Group='{SecurityElement.Escape(group.Trim())}'") +
                    "><Default></Default><Customization><ArrayOfProperty>" +
                    "<Property><Name>TextField</Name><Value xmlns:q6='http://www.w3.org/2001/XMLSchema' p4:type='q6:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>" +
                    noteId + "</Value></Property>" +
                    "</ArrayOfProperty></Customization></Field>";
        return (note, field);
    }

    /// <summary>
    /// Terms in a term set by path: found (loading each level once) or made.
    /// Terms made here are known without asking SharePoint again.
    /// </summary>
    private sealed class TermTree(ClientContext ctx, TermSet set, int lcid)
    {
        private readonly Dictionary<TermSetItem, List<Term>> children = [];

        public async Task<Term?> FindAsync(string[] path)
        {
            TermSetItem parent = set;
            Term? term = null;
            foreach (var part in path)
            {
                term = (await ChildrenAsync(parent).ConfigureAwait(false)).FirstOrDefault(t => SameTerm(t.Name, part));
                if (term is null)
                {
                    return null;
                }
                parent = term;
            }
            return term;
        }

        /// <summary>The term at the path, making it (and those above it) where missing; how many were made.</summary>
        public async Task<(Term Term, int Made)> EnsureAsync(string[] path)
        {
            TermSetItem parent = set;
            Term? term = null;
            var made = 0;
            foreach (var part in path)
            {
                var level = await ChildrenAsync(parent).ConfigureAwait(false);
                term = level.FirstOrDefault(t => SameTerm(t.Name, part));
                if (term is null)
                {
                    term = parent.CreateTerm(part, lcid, Guid.NewGuid());
                    term.Name = part;
                    level.Add(term);
                    children[term] = [];
                    made++;
                }
                parent = term;
            }
            return (term!, made);
        }

        private async Task<List<Term>> ChildrenAsync(TermSetItem parent)
        {
            if (!children.TryGetValue(parent, out var level))
            {
                ctx.Load(parent.Terms, ts => ts.Include(t => t.Name, t => t.Id));
                await ctx.RunAsync().ConfigureAwait(false);
                level = parent.Terms.ToList();
                children[parent] = level;
            }
            return level;
        }
    }
}
