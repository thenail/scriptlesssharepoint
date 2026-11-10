using System.Collections.ObjectModel;
using System.Globalization;

using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

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
        // every site collection again: the connection, not a saved selection, is the one picked
        if (active is not null)
        {
            SelectNavQuietly(active.Id);
        }
        ShowAllButton.Visibility = Visibility.Collapsed;
        ResetTable();
        query = string.IsNullOrWhiteSpace(SearchBox.Text) ? null : SearchBox.Text.Trim();
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
        // a saved selection opens again; anything else reloads the site collections
        if (shownSaved is { } s && savedStore.Load(s.File) is { } fresh)
        {
            _ = ShowSavedAsync(fresh);
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
        if (row.IsChecked && node.Kind != pickedKind)
        {
            pickedKind = node.Kind;
        }
        UpdatePane();
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
        pickedKind = NodeKind.Tenant;
        foreach (var r in rows)
        {
            r.SetChecked(false);
        }
        ShowOverview();
        UpdatePane();
    }

    private void OnClearSelection(object sender, RoutedEventArgs e) => ResetSelection();

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
        }
        UpdatePane();
    }

    // ---- the table's own states ----

    private void ResetTable()
    {
        rows.Clear();
        HideEmpty();
    }

    private void ShowEmpty(string text, bool ring = false)
    {
        rows.Clear();
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
