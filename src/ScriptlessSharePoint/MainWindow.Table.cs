using System.Collections.ObjectModel;
using System.Globalization;

using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.App;

/// <summary>
/// The table: the site collections, each opening in place to its subsites,
/// lists and libraries, and those to their folders, files and items. Ticks
/// are kept by what a row is (<see cref="SpNode.Key"/>), so they survive
/// closing and opening rows.
/// </summary>
public sealed partial class MainWindow
{
    /// <summary>How many site collections a page brings.</summary>
    private const int SitePageSize = 100;

    /// <summary>How long typing pauses before the search goes to SharePoint.</summary>
    private static readonly TimeSpan TypingPause = TimeSpan.FromMilliseconds(450);

    private readonly ObservableCollection<Row> rows = [];

    /// <summary>What's ticked, in the order it was ticked.</summary>
    private readonly OrderedDictionary<string, SpNode> selection = new();

    /// <summary>Bumped whenever what the table shows starts again, so late answers are dropped.</summary>
    private int view;

    /// <summary>Whether the admin centre's list of every site works for this connection; else search.</summary>
    private bool adminScope = true;

    private string? query;
    private DispatcherTimer? typing;

    // ---- site collections ----

    private async Task LoadSitesAsync()
    {
        var v = ++view;
        failedMore.Clear();
        shownSaved = null;
        shownPreset = null;
        shownRule = null;
        // every site collection again: the connection, not a saved selection, is the one picked
        if (active is not null)
        {
            SelectNavQuietly(active.Id);
        }
        ResetTable();
        query = string.IsNullOrWhiteSpace(SearchBox.Text) ? null : SearchBox.Text.Trim();
        SetCrumbs(query is null ? null : $"Search: {query}");
        ShowEmpty(query is null ? "Loading the site collections…" : $"Looking for “{query}”…", ring: true);
        SetStatus("Loading…", "", working: true);
        try
        {
            var page = await SitesPageAsync(null);
            if (v != view)
            {
                return;
            }
            AddSites(page, at: 0, parent: null);
            var n = page.Total ?? page.Sites.Count;
            var one = n == 1 && !page.TotalIsEstimate && (page.Total is not null || page.Next is null);
            var count = page.Total is not null
                ? $"{(page.TotalIsEstimate ? "About " : "")}{n:N0} site collection{(one ? "" : "s")}"
                : $"{n:N0}{(page.Next is null ? "" : "+")} site collection{(one ? "" : "s")}";
            SetStatus(query is null ? count : $"{count} {(one ? "matches" : "match")}", adminScope
                ? "Every site collection in the tenant, from the admin centre."
                : "The site collections you can open, from SharePoint search (not an administrator). New sites take a few minutes to show up.");
            if (rows.Count == 0)
            {
                ShowEmpty(query is null ? "There are no site collections to show." : $"No site collection matches “{query}”.");
            }
            else
            {
                HideEmpty();
            }
        }
        catch (Exception e)
        {
            if (v != view)
            {
                return;
            }
            ShowError(e);
            SetStatus("Couldn't load the site collections", "");
            ShowEmpty("Couldn't load the site collections.");
        }
    }

    /// <summary>A page of site collections: from the admin centre, or from search when that's refused.</summary>
    private async Task<SitePage> SitesPageAsync(string? cursor)
    {
        if (Demo.On)
        {
            return await Demo.SitesAsync(query);
        }
        var client = workspace.Current().Client;
        if (adminScope)
        {
            try
            {
                return await client.ListSitesAsAdminAsync(query, SitePageSize, cursor);
            }
            catch (Exception e) when (SharePointClient.IsAccessDenied(e) && cursor is null)
            {
                adminScope = false;
            }
        }
        return await client.SearchSitesAsync(query, SitePageSize, cursor);
    }

    private void AddSites(SitePage page, int at, Row? parent)
    {
        foreach (var site in page.Sites)
        {
            var node = SpNode.FromSite(site);
            rows.Insert(at++, Row.ForNode(this, node, parent, selection.ContainsKey(node.Key)));
        }
        if (page.Next is { } next)
        {
            rows.Insert(at, Row.More(this, parent, next, "Show more site collections"));
        }
    }

    private void OnSearchChanged(AutoSuggestBox sender, AutoSuggestBoxTextChangedEventArgs args)
    {
        if (args.Reason != AutoSuggestionBoxTextChangeReason.UserInput || identity is null)
        {
            return;
        }
        typing?.Stop();
        typing = new DispatcherTimer { Interval = TypingPause };
        typing.Tick += (_, _) =>
        {
            typing.Stop();
            if ((string.IsNullOrWhiteSpace(sender.Text) ? null : sender.Text.Trim()) != query)
            {
                _ = LoadSitesAsync();
            }
        };
        typing.Start();
    }

    private void OnSearchSubmitted(AutoSuggestBox sender, AutoSuggestBoxQuerySubmittedEventArgs args)
    {
        typing?.Stop();
        if (identity is not null)
        {
            _ = LoadSitesAsync();
        }
    }

    private void OnRefresh(object sender, RoutedEventArgs e)
    {
        if (TenantPage.Visibility != Visibility.Visible || identity is null)
        {
            return;
        }
        // a saved selection opens again; a predefined one (or a saved rule) reads the site collections again;
        // anything else reloads them (a plan on show stays)
        if (shownPreset is { } preset)
        {
            _ = ShowPresetAsync(preset, refresh: true, shownRule);
        }
        // (one from the history has no file: it shows again as it was)
        else if (shownSaved is { } s && (s.File.Length == 0 ? s : savedStore.Load(s.File)) is { } fresh)
        {
            ShowSaved(fresh);
        }
        else
        {
            _ = LoadSitesAsync();
        }
    }

    private void OnCollapseAll(object sender, RoutedEventArgs e)
    {
        for (var i = rows.Count - 1; i >= 0; i--)
        {
            if (rows[i].Depth > 0)
            {
                rows.RemoveAt(i);
            }
            else
            {
                rows[i].IsExpanded = false;
            }
        }
    }

    // ---- opening rows ----

    public async void Toggle(Row row)
    {
        if (row.IsLoading || row.Node is not { CanExpand: true })
        {
            return;
        }
        if (row.IsExpanded)
        {
            Collapse(row);
            return;
        }
        await ExpandAsync(row);
    }

    /// <summary>Opens a row: what's in it goes in below it (or a note: empty, an error).</summary>
    private async Task ExpandAsync(Row row)
    {
        if (row.IsExpanded || row.IsLoading || row.Node is not { CanExpand: true } node)
        {
            return;
        }
        var v = view;
        row.IsExpanded = true;
        row.IsLoading = true;
        try
        {
            IReadOnlyList<SpNode> children;
            string? next = null;
            if (node.Kind is NodeKind.SiteCollection or NodeKind.Subsite)
            {
                children = Demo.On ? await Demo.WebAsync(node) : await workspace.Current().Client.GetWebContentsAsync(node);
            }
            else
            {
                var page = Demo.On ? await Demo.FolderAsync(node) : await workspace.Current().Client.GetFolderContentsAsync(node);
                children = page.Nodes;
                next = page.Next;
            }
            if (v != view || !row.IsExpanded || !rows.Contains(row))
            {
                return;
            }
            var at = rows.IndexOf(row) + 1;
            if (children.Count == 0 && next is null)
            {
                rows.Insert(at, Row.Note(this, row, node.Kind is NodeKind.SiteCollection or NodeKind.Subsite
                    ? "No subsites, lists or libraries you can see."
                    : "Empty."));
            }
            InsertChildren(row, children, next, at);
        }
        catch (Exception e)
        {
            if (v == view && row.IsExpanded && rows.Contains(row))
            {
                rows.Insert(rows.IndexOf(row) + 1, Row.Note(this, row, SharePointErrors.Describe(e)));
            }
        }
        finally
        {
            row.IsLoading = false;
        }
    }

    private void InsertChildren(Row parent, IReadOnlyList<SpNode> children, string? next, int at)
    {
        foreach (var child in children)
        {
            rows.Insert(at++, Row.ForNode(this, child, parent, selection.ContainsKey(child.Key)));
        }
        if (next is not null)
        {
            rows.Insert(at, Row.More(this, parent, next, $"Show the next {SharePointClient.FolderPageSize:N0}"));
        }
    }

    /// <summary>The rows directly under an open row, "show more" included.</summary>
    private IEnumerable<Row> Children(Row row)
    {
        var i = rows.IndexOf(row);
        for (var j = i + 1; i >= 0 && j < rows.Count && rows[j].Depth > row.Depth; j++)
        {
            if (rows[j].Depth == row.Depth + 1)
            {
                yield return rows[j];
            }
        }
    }

    private void Collapse(Row row)
    {
        var i = rows.IndexOf(row);
        if (i < 0)
        {
            return;
        }
        while (i + 1 < rows.Count && rows[i + 1].Depth > row.Depth)
        {
            rows.RemoveAt(i + 1);
        }
        row.IsExpanded = false;
    }

    public async void LoadMore(Row more)
    {
        failedMore.Remove(more);
        await LoadMoreAsync(more);
    }

    /// <summary>
    /// The next page of site collections comes by itself when the end of the
    /// list scrolls into view: the ListView makes a row's container just before
    /// it shows. Its link stays for when a page fails. Folders' pages (500 each)
    /// stay a click, so a big library doesn't pour in.
    /// </summary>
    private void OnRowShown(ListViewBase sender, ContainerContentChangingEventArgs args)
    {
        if (!args.InRecycleQueue && args.Item is Row { IsMore: true, Parent: null } more && !failedMore.Contains(more))
        {
            _ = LoadMoreAsync(more);
        }
    }

    /// <summary>"Show more" rows whose page failed: they wait for a click, not for the next scroll.</summary>
    private readonly HashSet<Row> failedMore = [];

    private async Task LoadMoreAsync(Row more)
    {
        if (more.Cursor is not { } cursor || more.Message.EndsWith('…'))
        {
            return;
        }
        var v = view;
        var text = more.Message;
        more.Message = "Loading…";
        try
        {
            if (more.Parent is null)
            {
                var page = await SitesPageAsync(cursor);
                if (v != view || !rows.Contains(more))
                {
                    return;
                }
                var at = rows.IndexOf(more);
                rows.Remove(more);
                AddSites(page, at, null);
            }
            else
            {
                var page = Demo.On ? new NodePage([], null) : await workspace.Current().Client.GetFolderContentsAsync(more.Parent.Node!, cursor);
                if (v != view || !rows.Contains(more))
                {
                    return;
                }
                var at = rows.IndexOf(more);
                rows.Remove(more);
                InsertChildren(more.Parent, page.Nodes, page.Next, at);
            }
        }
        catch (Exception e)
        {
            more.Message = text;
            failedMore.Add(more);
            ShowError(e);
        }
    }

    // ---- ticks ----

    public void Checked(Row row)
    {
        if (row.Node is not { } node)
        {
            return;
        }
        Note(row);
        if (row.IsChecked && node.Kind != pickedKind)
        {
            pickedKind = node.Kind;
        }
        UpdatePane();
    }

    /// <summary>Ticks or unticks rows together, then brings the pane up to date once.</summary>
    private void TickRows(IEnumerable<Row> rowsToTick, bool on)
    {
        var changed = rowsToTick.Where(r => r.IsNode && r.IsChecked != on).ToList();
        if (changed.Count == 0)
        {
            return;
        }
        foreach (var r in changed)
        {
            r.SetChecked(on);
            Note(r);
        }
        if (on)
        {
            pickedKind = changed.GroupBy(r => r.Node!.Kind).OrderByDescending(g => g.Count()).First().Key;
        }
        UpdatePane();
    }

    /// <summary>Keeps a row's tick in the selection, with its site collection and parent (for saving), and on its twins.</summary>
    private void Note(Row row)
    {
        var node = row.Node!;
        if (row.IsChecked)
        {
            selection[node.Key] = node;
            var root = row;
            while (root.Parent is not null)
            {
                root = root.Parent;
            }
            if (root.Node is { Kind: NodeKind.SiteCollection } site)
            {
                siteOf[node.Key] = site;
            }
            if (row.Parent?.Node is { Kind: not NodeKind.SiteCollection } parent)
            {
                parentOf[node.Key] = parent;
            }
        }
        else
        {
            selection.Remove(node.Key);
        }
        // the same thing can show twice (e.g. a site found again): tick both
        foreach (var other in rows.Where(r => r != row && r.Node?.Key == node.Key))
        {
            other.SetChecked(row.IsChecked);
        }
    }

    // ---- a row's menu, and the keyboard ----

    /// <summary>The row an element of the table is in (its container, and the row).</summary>
    private (ListViewItem Container, Row Row)? RowAt(DependencyObject? element)
    {
        for (var e = element; e is not null && e != Table; e = Microsoft.UI.Xaml.Media.VisualTreeHelper.GetParent(e))
        {
            if (e is ListViewItem item && Table.ItemFromContainer(item) is Row row)
            {
                return (item, row);
            }
        }
        return null;
    }

    /// <summary>The rows beside a row: under the same open row, or the site collections.</summary>
    private List<Row> Siblings(Row row) => row.Parent is { } parent
        ? Children(parent).Where(r => r.IsNode).ToList()
        : rows.Where(r => r.Depth == 0 && r.IsNode).ToList();

    private void OnRowContextRequested(UIElement sender, ContextRequestedEventArgs e)
    {
        if (RowAt(e.OriginalSource as DependencyObject) is not ({ } container, { Node: { } node } row))
        {
            return;
        }
        e.Handled = true;
        var menu = new MenuFlyout();
        void Add(string text, string glyph, Action act, string? keys = null)
        {
            var item = new MenuFlyoutItem { Text = text, Icon = new FontIcon { Glyph = glyph } };
            if (keys is not null)
            {
                item.KeyboardAcceleratorTextOverride = keys;
            }
            item.Click += (_, _) => act();
            menu.Items.Add(item);
        }
        if (row.CanExpand)
        {
            Add(row.IsExpanded ? "Close" : "Open", row.IsExpanded ? "" : "", () => Toggle(row), row.IsExpanded ? "Left" : "Right");
        }
        Add("Open in the browser", "", () => OpenUrl(node.Url));
        Add("Copy the address", "", () => CopyText(node.Url));
        menu.Items.Add(new MenuFlyoutSeparator());
        Add(row.IsChecked ? "Untick" : "Tick", row.IsChecked ? "" : "", () => TickRows([row], !row.IsChecked), "Space");
        var siblings = Siblings(row);
        var where = row.Parent?.Node is { } parent ? $"in {parent.Title}" : "site collections shown";
        if (siblings.Any(r => !r.IsChecked))
        {
            Add($"Tick all {siblings.Count:N0} {where}", "", () => TickRows(siblings, true));
        }
        if (siblings.Count(r => r.IsChecked) > 1)
        {
            Add($"Untick all {where}", "", () => TickRows(siblings, false));
        }
        if (row.IsExpanded && Children(row).Where(r => r.IsNode).ToList() is { Count: > 0 } inside && inside.Any(r => !r.IsChecked))
        {
            Add($"Tick the {inside.Count:N0} in it", "", () => TickRows(inside, true));
        }
        if (e.TryGetPosition(container, out var at))
        {
            menu.ShowAt(container, at);
        }
        else
        {
            menu.ShowAt(container);
        }
    }

    /// <summary>
    /// Space ticks the row with the focus, Right opens it (or goes into it),
    /// Left closes it (or goes up to the row it's in), Enter opens or closes it.
    /// </summary>
    private void OnTableKeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (RowAt(FocusManager.GetFocusedElement(Table.XamlRoot) as DependencyObject) is not ({ } container, { Node: not null } row))
        {
            return;
        }
        var onRow = ReferenceEquals(FocusManager.GetFocusedElement(Table.XamlRoot), container);
        switch (e.Key)
        {
            case Windows.System.VirtualKey.Space when onRow:
                TickRows([row], !row.IsChecked);
                e.Handled = true;
                break;
            case Windows.System.VirtualKey.Enter when row.CanExpand:
                Toggle(row);
                e.Handled = true;
                break;
            case Windows.System.VirtualKey.Right when row.CanExpand:
                if (!row.IsExpanded)
                {
                    Toggle(row);
                }
                else if (Children(row).FirstOrDefault() is { } child)
                {
                    FocusRow(child);
                }
                e.Handled = true;
                break;
            case Windows.System.VirtualKey.Left:
                if (row.IsExpanded)
                {
                    Collapse(row);
                }
                else if (row.Parent is { } parent)
                {
                    FocusRow(parent);
                }
                e.Handled = true;
                break;
        }
    }

    private void FocusRow(Row row)
    {
        Table.ScrollIntoView(row);
        Table.UpdateLayout();
        (Table.ContainerFromItem(row) as ListViewItem)?.Focus(FocusState.Keyboard);
    }

    private void Untick(string key)
    {
        selection.Remove(key);
        foreach (var r in rows.Where(r => r.Node?.Key == key))
        {
            r.SetChecked(false);
        }
    }

    private void ResetSelection()
    {
        selection.Clear();
        siteOf.Clear();
        parentOf.Clear();
        pickedKind = NodeKind.Tenant;
        foreach (var r in rows)
        {
            r.SetChecked(false);
        }
        ShowOverview();
        UpdatePane();
    }

    private void OnClearSelection(object sender, RoutedEventArgs e)
    {
        KeepTicksBefore("Cleared the ticks", always: true);
        ResetSelection();
    }

    // ---- after a command ----

    /// <summary>Shows in the table what a command changed: renamed, gone, new contents, new sites.</summary>
    private void Apply(Command command, IReadOnlyList<TargetResult> results)
    {
        var changed = results.Where(r => r.Result is { Ok: true, WasSkipped: false }).ToList();
        switch (command.After)
        {
            case After.Update:
                foreach (var (target, result) in changed)
                {
                    if (result.Updated is not { } updated)
                    {
                        continue;
                    }
                    foreach (var row in rows.Where(r => r.Node?.Key == target.Key).ToList())
                    {
                        if (row.IsExpanded)
                        {
                            Collapse(row);
                        }
                        row.Node = updated;
                    }
                    if (selection.Remove(target.Key))
                    {
                        selection[updated.Key] = updated;
                    }
                }
                break;
            case After.Remove:
                foreach (var (target, _) in changed)
                {
                    foreach (var row in rows.Where(r => r.Node?.Key == target.Key).ToList())
                    {
                        Collapse(row);
                        rows.Remove(row);
                    }
                    selection.Remove(target.Key);
                }
                break;
            case After.ReloadChildren:
                foreach (var (target, _) in changed)
                {
                    foreach (var row in rows.Where(r => r.Node?.Key == target.Key).ToList())
                    {
                        if (row.IsExpanded)
                        {
                            Collapse(row);
                        }
                        Toggle(row);
                    }
                }
                break;
            case After.ReloadSites when changed.Count > 0:
                _ = LoadSitesAsync();
                break;
            case After.Tick:
                TickFound(changed);
                break;
        }
        UpdatePane();
    }

    /// <summary>
    /// What finders found becomes the selection, in place of the sites they
    /// looked in, each kept with its site and the row it's under (for saving).
    /// Nothing found: the selection stays as it was.
    /// </summary>
    private void TickFound(IReadOnlyList<TargetResult> results)
    {
        var found = results
            .SelectMany(r => r.Result.Found.Select(f => (f.Node, f.Parent, Site: r.Target.Kind == NodeKind.SiteCollection
                ? r.Target
                : siteOf.GetValueOrDefault(r.Target.Key) ?? SpNode.SiteFromUrl(r.Target.SiteUrl ?? r.Target.Url))))
            .DistinctBy(f => f.Node.Key)
            .ToList();
        if (found.Count == 0)
        {
            return;
        }
        KeepTicksBefore($"Ticked {Summary(found.Select(f => f.Node))}", always: true);
        foreach (var key in selection.Keys.ToList())
        {
            Untick(key);
        }
        siteOf.Clear();
        parentOf.Clear();
        foreach (var (node, parent, site) in found)
        {
            selection[node.Key] = node;
            siteOf[node.Key] = site;
            if (parent is { Kind: not NodeKind.SiteCollection })
            {
                parentOf[node.Key] = parent;
            }
            foreach (var row in rows.Where(r => r.Node?.Key == node.Key))
            {
                row.SetChecked(true);
            }
        }
        pickedKind = found.GroupBy(f => f.Node.Kind).OrderByDescending(g => g.Count()).ThenBy(g => g.Key).First().Key;
        if (ShowFound(results) is { } first)
        {
            Table.ScrollIntoView(first);
        }
    }

    /// <summary>
    /// Shows what finders found in the table: each site collection they looked
    /// in opens to just those, under the subsites, lists and folders they're
    /// in (nothing is loaded). Closing and opening a row shows all of it again.
    /// Returns the first found row.
    /// </summary>
    private Row? ShowFound(IReadOnlyList<TargetResult> results)
    {
        // by the site collection row each target is under: the rows down to it, and what's in it
        var groups = new Dictionary<Row, List<SpNode>>();
        foreach (var (target, result) in results.Where(r => r.Result.Found.Count > 0))
        {
            List<SpNode> chain;
            Row top;
            if (rows.FirstOrDefault(r => r.Node?.Key == target.Key) is { } row)
            {
                (top, chain) = (row, []);
                for (var r = row; r is not null; r = r.Parent)
                {
                    chain.Insert(0, r.Node!);
                    top = r;
                }
            }
            else
            {
                var site = siteOf.GetValueOrDefault(result.Found[0].Node.Key) ?? SpNode.SiteFromUrl(target.SiteUrl ?? target.Url);
                if (rows.FirstOrDefault(r => r.Depth == 0 && r.Node?.Key == site.Key) is not { } siteRow)
                {
                    continue;
                }
                (top, chain) = (siteRow, [site, target]);
            }
            var shown = groups.TryGetValue(top, out var list) ? list : groups[top] = [];
            shown.AddRange(chain);
            shown.AddRange(result.Found.Select(f => f.Parent).OfType<SpNode>());
            shown.AddRange(result.Found.Select(f => f.Node));
        }

        Row? first = null;
        foreach (var (top, shown) in groups)
        {
            Collapse(top);
            var row = InsertTree(shown.DistinctBy(n => n.Key).ToList(), top);
            if (row is not null && (first is null || rows.IndexOf(row) < rows.IndexOf(first)))
            {
                first = row;
            }
        }
        return first;
    }

    /// <summary>
    /// Puts rows in for <paramref name="shown"/>, each under the nearest of them
    /// above it (<see cref="SpNode.NearestAbove"/>), and keeps that as the
    /// ticked ones' parent. With <paramref name="root"/> (one of them, already
    /// in the table), they go under it; else at the end, site collections and
    /// whatever has nothing above it at the top. Returns the first ticked row.
    /// </summary>
    private Row? InsertTree(IReadOnlyList<SpNode> shown, Row? root)
    {
        var under = new Dictionary<string, List<SpNode>>();
        var top = new List<SpNode>();
        foreach (var node in shown)
        {
            if (node.Key == root?.Node?.Key)
            {
                continue;
            }
            var above = node.Kind == NodeKind.SiteCollection ? null : node.NearestAbove(shown);
            if (above is null && root is null)
            {
                top.Add(node);
                continue;
            }
            above ??= root!.Node!;
            (under.TryGetValue(above.Key, out var list) ? list : under[above.Key] = []).Add(node);
            if (above.Kind != NodeKind.SiteCollection && selection.ContainsKey(node.Key))
            {
                parentOf[node.Key] = above;
            }
        }

        var at = root is null ? rows.Count : rows.IndexOf(root) + 1;
        Row? first = null;
        void Add(SpNode node, Row? parent)
        {
            var row = Row.ForNode(this, node, parent, selection.ContainsKey(node.Key));
            rows.Insert(at++, row);
            if (row.IsChecked)
            {
                first ??= row;
            }
            AddUnder(node, row);
        }
        void AddUnder(SpNode node, Row row)
        {
            if (under.TryGetValue(node.Key, out var children))
            {
                row.IsExpanded = true;
                foreach (var child in children.OrderBy(n => n.CanExpand ? 0 : 1).ThenBy(n => n.Kind).ThenBy(n => n.Title, StringComparer.CurrentCultureIgnoreCase))
                {
                    Add(child, row);
                }
            }
        }
        if (root is not null)
        {
            AddUnder(root.Node!, root);
        }
        foreach (var node in top.OrderBy(n => n.Title, StringComparer.CurrentCultureIgnoreCase))
        {
            Add(node, null);
        }
        return first;
    }

    // ---- the table's own states ----

    private void ResetTable()
    {
        // what replaces the rows sets the title: a report on show just goes
        HideReport(restoreTitle: false);
        rows.Clear();
        HideEmpty();
    }

    private void ShowEmpty(string text, bool ring = false)
    {
        rows.Clear();
        SignInButton.Visibility = Visibility.Collapsed;
        EmptyText.Text = text;
        EmptyRing.IsActive = ring;
        EmptyRing.Visibility = ring ? Visibility.Visible : Visibility.Collapsed;
        EmptyState.Visibility = Visibility.Visible;
    }

    private void HideEmpty()
    {
        EmptyState.Visibility = Visibility.Collapsed;
        EmptyRing.IsActive = false;
    }

    private static string Plural(int n, string one, string many) =>
        $"{n.ToString("N0", CultureInfo.CurrentCulture)} {(n == 1 ? one : many)}";
}
