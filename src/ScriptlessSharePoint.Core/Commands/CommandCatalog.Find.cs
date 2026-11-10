using System.Text.RegularExpressions;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>
/// Finders: they look inside each ticked site, list or folder for something
/// (subsites, a library, the home page, what's at an address, files, folders
/// or items) and tick what they find instead, so the next command runs on
/// those. Read-only, run like any command, 4 targets at a time.
/// </summary>
public static partial class CommandCatalog
{
    /// <summary>The most webs one site's search goes through with its subsites.</summary>
    internal const int WebLimit = 200;

    /// <summary>The most rows one list's search reads (pages of 500).</summary>
    internal const int ScanLimit = 100_000;

    /// <summary>The most things one target ticks.</summary>
    internal const int FoundLimit = 5_000;

    private const string FindGroup = "Tick inside";

    /// <summary>
    /// Which lists to find. By template or by the name in the address, never by
    /// title: titles are translated ("Webstedssider"), templates and URLs aren't.
    /// </summary>
    private static readonly IReadOnlyList<Choice> ListsToFind =
    [
        new("sitepages", "Site Pages"),
        new("documents", "Documents (the site's default library)"),
        new("siteassets", "Site Assets"),
        new("pages", "Pages (publishing sites)"),
        new("style", "Style Library"),
        new("libraries", "Every document library"),
        new("lists", "Every list"),
        new("all", "Every list and library"),
        new("t:100", "Custom lists"),
        new("t:107,171", "Task lists"),
        new("t:106", "Calendars"),
        new("t:105", "Contacts lists"),
        new("t:104", "Announcements"),
        new("t:103", "Links lists"),
        new("t:1100", "Issue trackers"),
        new("t:102", "Surveys"),
        new("t:108", "Discussion boards"),
        new("t:109", "Picture libraries"),
        new("t:115", "Form libraries"),
        new("name", "The one with this name in its address"),
    ];

    private static readonly IReadOnlyList<Choice> ContentKinds =
    [
        new("files", "Files"),
        new("folders", "Folders"),
        new("items", "List items"),
        new("all", "Everything: files, folders and items"),
    ];

    private static CommandParameter WhichList(string label = "Which lists", string fallback = "sitepages", IReadOnlyList<NodeKind>? forKinds = null) =>
        new("which", label, ParameterKind.Choice) { Choices = ListsToFind, Default = fallback, ForKinds = forKinds };

    private static CommandParameter ListName(IReadOnlyList<NodeKind>? forKinds = null) =>
        new("list", "Name in the list's address", ParameterKind.Text)
        {
            Placeholder = "Contracts",
            Hint = "The last part of its address, as in /sites/hr/Contracts or /sites/hr/Lists/Contracts.",
            ShowIf = new ShowIf("which", "name"),
            ForKinds = forKinds,
        };

    private static CommandParameter NamePattern(string hint) =>
        new("name", "Name", ParameterKind.Text) { Required = false, Placeholder = "*.docx", Hint = hint };

    private static CommandParameter Subsites(string label = "Look in subsites too", IReadOnlyList<NodeKind>? forKinds = null) =>
        Toggle("subsites", label, false) with { ForKinds = forKinds };

    private static IEnumerable<Command> FindCommands()
    {
        yield return new Command
        {
            Id = "find.subsites",
            Title = "Tick the subsites in each",
            Group = FindGroup,
            Description = "Ticks each site's subsites instead of the site: the ones right below it, or every level down.",
            PowerShell = "Get-PnPSubWeb",
            Targets = Webs,
            ReadOnly = true,
            After = After.Tick,
            Glyph = "",
            Parameters =
            [
                Toggle("subsites", "Every level down", true, "Off: only the subsites right below each site."),
                NamePattern("Empty: all of them. Otherwise a word in the title or address, or a pattern with * and ?."),
            ],
            Run = async (c, n) =>
            {
                var pattern = c.Args["name"];
                var found = new List<FoundNode>();
                foreach (var (web, contents) in await WebsAsync(c, n, c.Args.Toggle("subsites")))
                {
                    found.AddRange(contents
                        .Where(x => x.Kind == NodeKind.Subsite && NameMatches(x, pattern))
                        .Select(x => new FoundNode(x, web.Kind == NodeKind.Subsite ? web : null)));
                }
                return FoundResult(found);
            },
        };

        yield return new Command
        {
            Id = "find.lists",
            Title = "Tick a library or list in each",
            Group = FindGroup,
            Description = "Looks in each site for libraries or lists of a kind, and ticks them instead of the sites, to run a library or list command on them all.",
            PowerShell = "Get-PnPList",
            Targets = Webs,
            ReadOnly = true,
            After = After.Tick,
            Glyph = "",
            Parameters = [WhichList(), ListName(), Subsites()],
            Run = async (c, n) =>
            {
                var found = new List<FoundNode>();
                foreach (var (list, web) in await ListsAsync(c, n))
                {
                    found.Add(new FoundNode(list, web.Kind == NodeKind.Subsite ? web : null));
                }
                return FoundResult(found);
            },
        };

        yield return new Command
        {
            Id = "find.home",
            Title = "Tick the home page of each",
            Group = FindGroup,
            Description = "Finds each site's home page (usually Home.aspx in Site Pages) and ticks it instead of the site, to run a file command on them all.",
            PowerShell = "Get-PnPHomePage",
            Targets = Webs,
            ReadOnly = true,
            After = After.Tick,
            Glyph = "",
            Parameters = [Subsites("Subsites' home pages too")],
            Run = async (c, n) =>
            {
                var found = new List<FoundNode>();
                foreach (var web in await WebsOnlyAsync(c, n))
                {
                    if (await c.Client.GetHomePageAsync(web, c.Cancellation) is { } home)
                    {
                        found.Add(home);
                    }
                }
                return FoundResult(found);
            },
        };

        yield return new Command
        {
            Id = "find.path",
            Title = "Tick what's at an address in each",
            Group = FindGroup,
            Description = "Ticks the same folder, file, list or library in each site, by its address from the site's own: a policy folder every site has, a page, a list.",
            PowerShell = "Get-PnPFolder / Get-PnPFile -Url",
            Targets = Webs,
            ReadOnly = true,
            After = After.Tick,
            Glyph = "",
            Parameters =
            [
                new("path", "Address in the site", ParameterKind.Text)
                {
                    Placeholder = "Shared Documents/Policies",
                    Hint = "After the site's own address, as in Shared Documents/Policies, SitePages/About.aspx or Lists/Contacts.",
                },
                Subsites(),
            ],
            Run = async (c, n) =>
            {
                var path = c.Args.Text("path");
                var found = new List<FoundNode>();
                foreach (var web in await WebsOnlyAsync(c, n))
                {
                    if (await c.Client.GetAtPathAsync(web, path, c.Cancellation) is { } at)
                    {
                        found.Add(at);
                    }
                }
                return FoundResult(found);
            },
        };

        yield return new Command
        {
            Id = "find.content",
            Title = "Tick files, folders or items in each",
            Group = FindGroup,
            Description = "Looks through each list, library or folder (or, for a site, the lists picked below) and ticks the files, folders or items that match, to run a command on them all.",
            PowerShell = "Get-PnPListItem / Find-PnPFile",
            Targets = [NodeKind.SiteCollection, NodeKind.Subsite, NodeKind.List, NodeKind.Library, NodeKind.Folder],
            ReadOnly = true,
            After = After.Tick,
            Glyph = "",
            Parameters =
            [
                new("what", "What", ParameterKind.Choice) { Choices = ContentKinds, Default = "files" },
                NamePattern("Empty: all of them. Otherwise a word in the name or title, or a pattern with * and ?, as in *.docx or Budget 202?.xlsx."),
                Toggle("subfolders", "Look in subfolders too", true),
                Toggle("checkedout", "Only files checked out to someone", false),
                WhichList("In which lists", "documents", Webs),
                ListName(Webs),
                Subsites(forKinds: Webs),
            ],
            Run = async (c, n) =>
            {
                var what = c.Args.Text("what");
                var pattern = c.Args["name"];
                var checkedOut = c.Args.Toggle("checkedout");
                var containers = n.Kind is NodeKind.SiteCollection or NodeKind.Subsite
                    ? (await ListsAsync(c, n)).Select(l => l.List).ToList()
                    : [n];
                var found = new List<FoundNode>();
                var stopped = false;
                foreach (var container in containers)
                {
                    stopped |= await ContentsAsync(c, container, x => ContentMatches(x, what, pattern, checkedOut), found);
                    if (found.Count >= FoundLimit)
                    {
                        break;
                    }
                }
                return FoundResult(found, stopped || found.Count >= FoundLimit
                    ? $"Stopped at {FoundLimit:N0} found or {ScanLimit:N0} looked through in a list: tick fewer, or narrow the name."
                    : null);
            },
        };
    }

    // ---- matching ----

    /// <summary>Whether a list is the one asked for (<see cref="ListsToFind"/>).</summary>
    internal static bool ListMatches(SpNode list, string which, string? name)
    {
        var leaf = list.LeafName;
        if (which.StartsWith("t:", StringComparison.Ordinal))
        {
            return which[2..].Split(',').Contains(list.Template);
        }
        return which switch
        {
            "sitepages" => list.Template == "119",
            "documents" => list.Kind == NodeKind.Library && Same(leaf, "Shared Documents"),
            "siteassets" => Same(leaf, "SiteAssets"),
            "pages" => list.Template == "850",
            "style" => Same(leaf, "Style Library"),
            "libraries" => list.Kind == NodeKind.Library,
            "lists" => list.Kind == NodeKind.List,
            "all" => list.Kind is NodeKind.List or NodeKind.Library,
            "name" => name?.Trim().Trim('/') is { Length: > 0 } n &&
                      (Same(leaf, n) || list.ServerRelativeUrl?.TrimEnd('/').EndsWith("/" + n, StringComparison.OrdinalIgnoreCase) == true),
            _ => false,
        };

        static bool Same(string a, string b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Whether a name or title fits: anything when there's no pattern; a
    /// pattern with * or ? has to match it whole; plain words just appear in it.
    /// </summary>
    internal static bool NameMatches(SpNode node, string? pattern)
    {
        if (string.IsNullOrWhiteSpace(pattern))
        {
            return true;
        }
        var p = pattern.Trim();
        string[] names = [node.LeafName, node.Title];
        if (p.Contains('*') || p.Contains('?'))
        {
            var regex = new Regex("^" + Regex.Escape(p).Replace(@"\*", ".*").Replace(@"\?", ".") + "$",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
            return names.Any(regex.IsMatch);
        }
        return names.Any(n => n.Contains(p, StringComparison.CurrentCultureIgnoreCase));
    }

    internal static bool ContentMatches(SpNode node, string what, string? pattern, bool checkedOut) =>
        what switch
        {
            "files" => node.Kind == NodeKind.File,
            "folders" => node.Kind == NodeKind.Folder,
            "items" => node.Kind == NodeKind.Item,
            _ => node.Kind is NodeKind.File or NodeKind.Folder or NodeKind.Item,
        } &&
        (!checkedOut || node.CheckedOutTo is not null) &&
        NameMatches(node, pattern);

    // ---- looking ----

    /// <summary>A web and what's in it, then (when asked) each subsite below it and what's in that.</summary>
    private static async Task<List<(SpNode Web, IReadOnlyList<SpNode> Contents)>> WebsAsync(CommandContext c, SpNode web, bool subsites)
    {
        var all = new List<(SpNode, IReadOnlyList<SpNode>)>();
        var queue = new Queue<SpNode>([web]);
        while (queue.TryDequeue(out var w) && all.Count < WebLimit)
        {
            c.Cancellation.ThrowIfCancellationRequested();
            var contents = await c.Client.GetWebContentsAsync(w, c.Cancellation);
            all.Add((w, contents));
            if (subsites)
            {
                foreach (var s in contents.Where(x => x.Kind == NodeKind.Subsite))
                {
                    queue.Enqueue(s);
                }
            }
        }
        return all;
    }

    /// <summary>The target, and its subsites when the "subsites" toggle is on.</summary>
    private static async Task<IReadOnlyList<SpNode>> WebsOnlyAsync(CommandContext c, SpNode web) =>
        c.Args.Toggle("subsites") ? (await WebsAsync(c, web, true)).Select(w => w.Web).ToList() : [web];

    /// <summary>The lists in a site (and its subsites, when asked) that "which" and "list" pick, each with its web.</summary>
    private static async Task<List<(SpNode List, SpNode Web)>> ListsAsync(CommandContext c, SpNode web)
    {
        var which = c.Args.Text("which");
        var name = c.Args["list"];
        if (which == "name" && name is null)
        {
            throw new ArgumentException("Name in the list's address: required for “the one with this name”.");
        }
        return (await WebsAsync(c, web, c.Args.Toggle("subsites")))
            .SelectMany(w => w.Contents
                .Where(x => x.Kind is NodeKind.List or NodeKind.Library && ListMatches(x, which, name))
                .Select(x => (x, w.Web)))
            .ToList();
    }

    /// <summary>
    /// Pages through a list, library or folder (and its subfolders, when asked),
    /// adding what matches to <paramref name="found"/>, each under the folder
    /// it's in when that's been seen, else under the container. True when it
    /// stopped before the end.
    /// </summary>
    private static async Task<bool> ContentsAsync(CommandContext c, SpNode container, Func<SpNode, bool> matches, List<FoundNode> found)
    {
        var folders = new Dictionary<string, SpNode>(StringComparer.OrdinalIgnoreCase);
        if (container.ServerRelativeUrl is { } root)
        {
            folders[root.TrimEnd('/')] = container;
        }
        string? cursor = null;
        var scanned = 0;
        do
        {
            c.Cancellation.ThrowIfCancellationRequested();
            var page = await c.Client.GetFolderContentsAsync(container, cursor, c.Args.Toggle("subfolders"), c.Cancellation);
            foreach (var x in page.Nodes)
            {
                if (x.Kind == NodeKind.Folder && x.ServerRelativeUrl is { } u)
                {
                    folders[u.TrimEnd('/')] = x;
                }
            }
            foreach (var x in page.Nodes.Where(matches))
            {
                var parent = x.ParentFolder is { } p && folders.TryGetValue(p.TrimEnd('/'), out var f) ? f : container;
                found.Add(new FoundNode(x, parent));
                if (found.Count >= FoundLimit)
                {
                    return true;
                }
            }
            scanned += page.Nodes.Count;
            cursor = page.Next;
        }
        while (cursor is not null && scanned < ScanLimit);
        return cursor is not null;
    }

    private static CommandResult FoundResult(List<FoundNode> found, string? note = null)
    {
        if (found.Count == 0)
        {
            return CommandResult.Skipped("None here.");
        }
        var names = string.Join(", ", found.Select(f => f.Node.Title).Take(5)) + (found.Count > 5 ? $" and {found.Count - 5:N0} more" : "");
        return CommandResult.Done(names, note) with { Found = found };
    }
}
