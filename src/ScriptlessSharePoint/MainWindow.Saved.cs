using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

using ScriptlessSharePoint.Selections;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.App;

/// <summary>
/// Selections saved under a name, as File Size has them: listed in the pane
/// below the connections, after a gap. Picking one makes it the selection and
/// shows it in its connection's table, ticked, with the sites, lists and
/// folders down to each item opened. Kept in Selections\ in the data folder,
/// one file each (Core\Selections).
/// </summary>
public sealed partial class MainWindow
{
    /// <summary>The navigation item's tag: this, then the saved selection's file.</summary>
    private const string SavedTag = "saved:";

    /// <summary>Rows opened to show one saved selection: enough for any sensible list, not so many the table drowns.</summary>
    private const int OpenLimit = 300;

    /// <summary>"Show more" pages followed in one folder while looking for an item.</summary>
    private const int PageLimit = 40;

    private readonly SavedSelectionStore savedStore = new();

    /// <summary>The site collection each ticked thing is in, for saving the selection.</summary>
    private readonly Dictionary<string, SpNode> siteOf = [];

    /// <summary>The saved selection the table shows, instead of every site collection.</summary>
    private SavedSelection? shownSaved;

    /// <summary>A saved selection to show once its connection is signed in.</summary>
    private SavedSelection? pendingSaved;

    private int opened;

    // ---- in the navigation pane ----

    /// <summary>
    /// The saved selections below the connections, for connections that still
    /// exist. The one picked in the pane stays picked; if it's gone, its
    /// connection is picked instead.
    /// </summary>
    private void RefreshSavedSelections()
    {
        var was = navigating;
        navigating = true;
        var picked = (Nav.SelectedItem as NavigationViewItem)?.Tag as string;
        try
        {
            foreach (var old in Nav.MenuItems.OfType<NavigationViewItem>().Where(i => i.Tag is string t && t.StartsWith(SavedTag, StringComparison.Ordinal)).ToList())
            {
                Nav.MenuItems.Remove(old);
            }
            var connections = workspace.Load().Connections.ToDictionary(c => c.Id);
            var first = true;
            foreach (var s in savedStore.List())
            {
                if (!connections.TryGetValue(s.ConnectionId, out var connection))
                {
                    continue;
                }
                var item = new NavigationViewItem
                {
                    Content = s.Name,
                    Tag = SavedTag + s.File,
                    Icon = new FontIcon { Glyph = "" },
                    // the gap between the connections and these
                    Margin = first ? new Thickness(0, 16, 0, 0) : new Thickness(0),
                    ContextFlyout = SavedMenu(s.File),
                };
                ToolTipService.SetToolTip(item, $"{Summary(s.Items)}\n{connection.Name} · saved {s.Saved.LocalDateTime:g}");
                Nav.MenuItems.Add(item);
                first = false;
            }
            if (picked?.StartsWith(SavedTag, StringComparison.Ordinal) == true)
            {
                Nav.SelectedItem = NavItem(picked) ?? (active is null ? null : NavItem(active.Id));
            }
        }
        finally
        {
            navigating = was;
        }
    }

    private NavigationViewItem? NavItem(string tag) =>
        Nav.MenuItems.OfType<NavigationViewItem>().FirstOrDefault(i => i.Tag as string == tag);

    /// <summary>"3 libraries · 12 files".</summary>
    private static string Summary(IEnumerable<SpNode> nodes) =>
        string.Join(" · ", nodes.GroupBy(n => n.Kind).OrderBy(g => g.Key).Select(g => NodeKinds.Count(g.Key, g.Count())));

    private MenuFlyout SavedMenu(string file)
    {
        var menu = new MenuFlyout();
        var rename = new MenuFlyoutItem { Text = "Rename…", Icon = new FontIcon { Glyph = "" } };
        rename.Click += async (_, _) => await RenameSavedAsync(file);
        menu.Items.Add(rename);
        var remove = new MenuFlyoutItem { Text = "Remove from the list", Icon = new FontIcon { Glyph = "" } };
        remove.Click += async (_, _) => await RemoveSavedAsync(file);
        menu.Items.Add(remove);
        return menu;
    }

    /// <summary>Picks a navigation item (a connection, or a saved selection) without showing it again.</summary>
    private void SelectNavQuietly(string tag)
    {
        var item = NavItem(tag);
        if (item is null || ReferenceEquals(Nav.SelectedItem, item))
        {
            return;
        }
        var was = navigating;
        navigating = true;
        try
        {
            Nav.SelectedItem = item;
        }
        finally
        {
            navigating = was;
        }
    }

    // ---- saving, renaming, removing ----

    private async void OnSaveSelection(object sender, RoutedEventArgs e) => await SaveSelectionAsync();

    private async Task SaveSelectionAsync()
    {
        if (selection.Count == 0 || active is null)
        {
            return;
        }
        var name = await AskNameAsync("Save the selection", $"Selection {DateTime.Now:yyyy-MM-dd HH:mm}", "Save");
        if (name is null || selection.Count == 0 || active is null)
        {
            return;
        }
        var items = selection.Values.ToList();
        var sites = items.Select(n => siteOf.TryGetValue(n.Key, out var site) ? site : SpNode.SiteFromUrl(n.SiteUrl ?? n.Url));
        try
        {
            savedStore.Save(name, active.Id, items, sites);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            ShowError(new InvalidOperationException($"Couldn't save the selection: {ex.Message}"));
            return;
        }
        RefreshSavedSelections();
    }

    private async Task RenameSavedAsync(string file)
    {
        if (savedStore.Load(file) is not { } s)
        {
            RefreshSavedSelections();
            return;
        }
        var name = await AskNameAsync("Rename the saved selection", s.Name, "Rename");
        if (name is null || name == s.Name)
        {
            return;
        }
        savedStore.Rename(file, name);
        if (shownSaved?.File == file)
        {
            shownSaved = shownSaved with { Name = name };
            StatusTitle.Text = $"Saved selection: {name}";
        }
        RefreshSavedSelections();
    }

    private async Task RemoveSavedAsync(string file)
    {
        if (savedStore.Load(file) is not { } s)
        {
            RefreshSavedSelections();
            return;
        }
        if (!await ConfirmAsync("Remove the saved selection?",
                $"“{s.Name}” goes from the list. What it lists stays as it is in SharePoint.", "Remove", danger: false))
        {
            return;
        }
        savedStore.Delete(file);
        RefreshSavedSelections();
    }

    /// <summary>A name, or null when cancelled. Save stays off while the name is empty.</summary>
    private async Task<string?> AskNameAsync(string title, string name, string action)
    {
        var box = new TextBox { Header = "Name", Text = name, MinWidth = 320 };
        var dialog = new ContentDialog
        {
            XamlRoot = Root.XamlRoot,
            Style = (Style)Application.Current.Resources["DefaultContentDialogStyle"],
            Title = title,
            Content = box,
            PrimaryButtonText = action,
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Primary,
        };
        box.TextChanged += (_, _) => dialog.IsPrimaryButtonEnabled = box.Text.Trim().Length > 0;
        dialog.Opened += (_, _) =>
        {
            box.Focus(FocusState.Programmatic);
            box.SelectAll();
        };
        return await dialog.ShowAsync() == ContentDialogResult.Primary && box.Text.Trim() is { Length: > 0 } n ? n : null;
    }

    // ---- opening one ----

    private async Task OpenSavedAsync(string file)
    {
        if (savedStore.Load(file) is not { } s)
        {
            // removed under us, by hand: out of the list, back to the connection
            RefreshSavedSelections();
            return;
        }
        if (workspace.Load().Find(s.ConnectionId) is null)
        {
            RefreshSavedSelections();
            return;
        }
        // it stays picked in the pane, while the table shows its connection's sites
        if (active?.Id != s.ConnectionId || identity is null)
        {
            pendingSaved = s;
            await ShowConnectionAsync(s.ConnectionId, force: active?.Id == s.ConnectionId);
            return;
        }
        await ShowSavedAsync(s);
    }

    /// <summary>After signing in: the saved selection waiting for it, or every site collection.</summary>
    private Task ShowSitesOrSavedAsync()
    {
        if (pendingSaved is { } s && s.ConnectionId == active?.Id)
        {
            pendingSaved = null;
            return ShowSavedAsync(s);
        }
        pendingSaved = null;
        return LoadSitesAsync();
    }

    /// <summary>
    /// The saved selection becomes the selection, and the table shows the site
    /// collections it's in, opened down to each item. What can't be found any
    /// more (moved, renamed, deleted, or not for this account) is left unticked.
    /// </summary>
    private async Task ShowSavedAsync(SavedSelection s)
    {
        var v = ++view;
        shownSaved = s;
        SelectNavQuietly(SavedTag + s.File);
        typing?.Stop();
        SearchBox.Text = "";
        query = null;
        ShowPage(TenantPage);
        HideError();

        selection.Clear();
        siteOf.Clear();
        var sites = s.Sites.ToDictionary(n => n.Key);
        foreach (var item in s.Items)
        {
            selection[item.Key] = item;
            var site = SpNode.SiteFromUrl(item.SiteUrl ?? item.Url);
            siteOf[item.Key] = sites.GetValueOrDefault(site.Key) ?? site;
        }

        ResetTable();
        foreach (var root in siteOf.Values.DistinctBy(n => n.Key).OrderBy(n => n.Title, StringComparer.CurrentCultureIgnoreCase))
        {
            rows.Add(Row.ForNode(this, root, null, selection.ContainsKey(root.Key)));
        }
        ShowPane(true);
        ShowOverview();
        UpdatePane();
        ShowAllButton.Visibility = Visibility.Visible;
        SetStatus($"Saved selection: {s.Name}", $"Opening the way down to {Plural(s.Items.Count, "item", "items")}…", working: true);

        // one at a time, in table order, so each opens on what the last one opened
        opened = 0;
        var missing = new List<SpNode>();
        Row? first = null;
        foreach (var item in s.Items.OrderBy(n => siteOf[n.Key].Title, StringComparer.CurrentCultureIgnoreCase).ThenBy(n => n.Url, StringComparer.OrdinalIgnoreCase))
        {
            var row = await RevealAsync(item, v);
            if (v != view)
            {
                return;
            }
            if (row is null)
            {
                missing.Add(item);
            }
            else
            {
                first ??= row;
            }
        }
        foreach (var m in missing)
        {
            Untick(m.Key);
        }
        // the pane acts on what was saved: the kind there's most of
        if (selection.Count > 0)
        {
            pickedKind = selection.Values.GroupBy(n => n.Kind).OrderByDescending(g => g.Count()).ThenBy(g => g.Key).First().Key;
        }
        UpdatePane();
        SetStatus($"Saved selection: {s.Name}", missing.Count == 0
            ? $"{Summary(s.Items)}, from {Plural(s.Sites.Count, "site collection", "site collections")}."
            : $"{Plural(missing.Count, "item wasn't", "items weren't")} found, so {(missing.Count == 1 ? "it isn't" : "they aren't")} ticked: moved, renamed or deleted, or not for this account.");
        if (first is not null)
        {
            Table.ScrollIntoView(first);
        }
    }

    /// <summary>
    /// Opens the rows on the way down to <paramref name="target"/> (its site,
    /// subsites, list, folders, following "show more" where it has to) and
    /// returns its row, or null when it isn't there.
    /// </summary>
    private async Task<Row?> RevealAsync(SpNode target, int v)
    {
        var row = rows.FirstOrDefault(r => r.Depth == 0 && r.Node is { } n && n.Leads(target));
        while (row?.Node is { } node && v == view)
        {
            if (node.Key == target.Key)
            {
                return row;
            }
            if (!node.CanExpand)
            {
                return null;
            }
            if (!row.IsExpanded)
            {
                if (opened++ >= OpenLimit)
                {
                    return null;
                }
                await ExpandAsync(row);
            }
            Row? next = null;
            for (var page = 0; page < PageLimit && v == view; page++)
            {
                next = Children(row).FirstOrDefault(c => c.Node is { } n && n.Leads(target));
                if (next is not null || Children(row).FirstOrDefault(c => c.IsMore) is not { } more)
                {
                    break;
                }
                await LoadMoreAsync(more);
            }
            row = next;
        }
        return null;
    }

    private void OnShowAll(object sender, RoutedEventArgs e)
    {
        if (identity is not null)
        {
            _ = LoadSitesAsync();
        }
    }
}
