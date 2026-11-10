using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;

using Windows.ApplicationModel.DataTransfer;

using ScriptlessSharePoint.Selections;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.App;

/// <summary>
/// Selections saved under a name, as File Size has them: listed in the pane
/// below the connections, after a gap. Picking one makes it the selection and
/// shows it in its connection's table, ticked, each item under the row it was
/// ticked in, without loading anything around it. Kept in Selections\ in the
/// data folder, one file each (Core\Selections).
/// </summary>
public sealed partial class MainWindow
{
    /// <summary>The navigation item's tag: this, then the saved selection's file.</summary>
    private const string SavedTag = "saved:";

    private readonly SavedSelectionStore savedStore = new();

    /// <summary>The site collection each ticked thing is in, for saving the selection.</summary>
    private readonly Dictionary<string, SpNode> siteOf = [];

    /// <summary>The subsite, list or folder each ticked thing was ticked in, for saving the selection.</summary>
    private readonly Dictionary<string, SpNode> parentOf = [];

    /// <summary>The saved selection the table shows, instead of every site collection.</summary>
    private SavedSelection? shownSaved;

    /// <summary>A saved selection to show once its connection is signed in.</summary>
    private SavedSelection? pendingSaved;

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
                    AllowDrop = true,
                    // how many items it has
                    InfoBadge = new InfoBadge { Value = s.Items.Count },
                };
                var file = s.File;
                item.DragOver += (_, e) => MarkDropTarget(item, OnSavedDragOver(file, e));
                item.DragLeave += (_, _) => MarkDropTarget(item, false);
                item.Drop += (_, _) =>
                {
                    MarkDropTarget(item, false);
                    OnSavedDrop(file);
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

    /// <summary>A navigation item by its tag, the predefined selections inside their section included.</summary>
    private NavigationViewItem? NavItem(string tag) => NavItems(Nav.MenuItems).FirstOrDefault(i => i.Tag as string == tag);

    private static IEnumerable<NavigationViewItem> NavItems(IList<object> items) =>
        items.OfType<NavigationViewItem>().SelectMany(i => NavItems(i.MenuItems).Prepend(i));

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

    private async void OnSaveSelection(SplitButton sender, SplitButtonClickEventArgs e) => await SaveSelectionAsync();

    private async void OnSaveSelectionAsNew(object sender, RoutedEventArgs e) => await SaveSelectionAsync(asNew: true);

    /// <summary>The saved selection on show, when there is one of this connection: Save goes to it.</summary>
    private SavedSelection? SaveTarget =>
        shownSaved is { File.Length: > 0 } s && s.ConnectionId == active?.Id && File.Exists(s.File) ? s : null;

    /// <summary>Save's tooltip: where it saves to.</summary>
    private void UpdateSaveButton() => ToolTipService.SetToolTip(SaveSelectionButton,
        SaveTarget is { } s ? $"Save the changes to “{s.Name}” (Ctrl+S)" : "Save the selection (Ctrl+S)");

    /// <summary>
    /// Saves what's ticked: into the saved selection on show (what's ticked now
    /// is what it holds), or, with <paramref name="asNew"/> or none on show, as
    /// a new one under a name.
    /// </summary>
    private async Task SaveSelectionAsync(bool asNew = false)
    {
        if (selection.Count == 0 || active is null)
        {
            return;
        }
        if (!asNew && SaveTarget is { } target)
        {
            var items = selection.Values.ToList();
            SavedSelection? updated;
            try
            {
                updated = savedStore.Update(target.File, items,
                    items.Select(n => siteOf.TryGetValue(n.Key, out var site) ? site : SpNode.SiteFromUrl(n.SiteUrl ?? n.Url)),
                    items.Select(n => parentOf.GetValueOrDefault(n.Key)).OfType<SpNode>());
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                ShowError(new InvalidOperationException($"Couldn't save the selection: {ex.Message}"));
                return;
            }
            if (updated is not null)
            {
                var before = target.Items.Select(n => n.Key).ToHashSet();
                var now = updated.Items.Select(n => n.Key).ToHashSet();
                var added = now.Count(k => !before.Contains(k));
                var removed = before.Count(k => !now.Contains(k));
                shownSaved = updated;
                ShowToast($"Saved “{updated.Name}”", string.Join(" · ", new[]
                {
                    added > 0 ? $"{added:N0} added" : null,
                    removed > 0 ? $"{removed:N0} removed" : null,
                    added + removed == 0 ? "No changes" : null,
                    $"{Plural(updated.Items.Count, "item", "items")} in all",
                }.OfType<string>()));
                RefreshSavedSelections();
                return;
            }
        }
        await SaveAsNewAsync();
    }

    private async Task SaveAsNewAsync()
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
        SavedSelection saved;
        try
        {
            saved = savedStore.Save(name, active.Id, items, sites, items.Select(n => parentOf.GetValueOrDefault(n.Key)).OfType<SpNode>());
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            ShowError(new InvalidOperationException($"Couldn't save the selection: {ex.Message}"));
            return;
        }
        RefreshSavedSelections();
        // saved from another saved selection (or a predefined one): the new one is the one on show now, and Save goes to it
        if (shownSaved is not null)
        {
            shownSaved = saved;
            shownPreset = null;
            SelectNavQuietly(SavedTag + saved.File);
            StatusTitle.Text = $"Saved selection: {saved.Name}";
        }
        UpdateSaveButton();
        ShowToast($"Saved “{saved.Name}”", $"{Plural(saved.Items.Count, "item", "items")} · in the list on the left",
            shownSaved?.File == saved.File ? null : ("Open", () => Nav.SelectedItem = NavItem(SavedTag + saved.File)));
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
        ShowSaved(s);
    }

    /// <summary>After signing in: the saved selection waiting for it, or every site collection.</summary>
    private Task ShowSitesOrSavedAsync()
    {
        if (pendingPlan is { } plan && plan.ConnectionId == active?.Id)
        {
            pendingPlan = null;
            pendingPreset = null;
            pendingSaved = null;
            return ShowPlanSelectionAsync(plan);
        }
        if (pendingPreset is { } preset)
        {
            pendingPreset = null;
            pendingSaved = null;
            return ShowPresetAsync(preset, refresh: false);
        }
        if (pendingSaved is null)
        {
            pendingForm = null;
        }
        if (pendingSaved is { } s && s.ConnectionId == active?.Id)
        {
            pendingSaved = null;
            ShowSaved(s);
            return Task.CompletedTask;
        }
        pendingSaved = null;
        return LoadSitesAsync();
    }

    /// <summary>
    /// The saved selection becomes the selection, and the table shows the site
    /// collections it's in, each item under the subsite, list or folder it was
    /// ticked in. Nothing is loaded: those rows show only what was saved until
    /// they're closed and opened again.
    /// </summary>
    private void ShowSaved(SavedSelection s)
    {
        ++view;
        shownSaved = s;
        shownPreset = null;
        SelectNavQuietly(SavedTag + s.File);
        typing?.Stop();
        SearchBox.Text = "";
        query = null;
        ShowPage(TenantPage);
        HideError();

        selection.Clear();
        siteOf.Clear();
        parentOf.Clear();
        var sites = s.Sites.ToDictionary(n => n.Key);
        foreach (var item in s.Items)
        {
            selection[item.Key] = item;
            var site = SpNode.SiteFromUrl(item.SiteUrl ?? item.Url);
            siteOf[item.Key] = sites.GetValueOrDefault(site.Key) ?? site;
        }

        // every row there is to show, each under the nearest one above it
        var roots = siteOf.Values.DistinctBy(n => n.Key).ToList();
        ResetTable();
        var first = InsertTree(roots.Concat(s.Parents).Concat(s.Items).DistinctBy(n => n.Key).ToList(), null);

        // the pane acts on what was saved: the kind there's most of
        if (selection.Count > 0)
        {
            pickedKind = selection.Values.GroupBy(n => n.Kind).OrderByDescending(g => g.Count()).ThenBy(g => g.Key).First().Key;
        }
        ShowPane(true);
        ShowOverview();
        UpdatePane();
        ShowAllButton.Visibility = Visibility.Visible;
        SetStatus(s.File.Length == 0 ? s.Name : $"Saved selection: {s.Name}", $"{Summary(s.Items)}, from {Plural(roots.Count, "site collection", "site collections")}.");
        if (first is not null)
        {
            Table.ScrollIntoView(first);
        }
        OpenPendingForm();
    }

    // ---- dragging rows onto one ----

    /// <summary>What a drag from the table carries: each item, its site collection, and the row it's under.</summary>
    private sealed record Dragged(SpNode Node, SpNode Site, SpNode? Parent);

    /// <summary>The rows being dragged, while a drag from the table is on (it never leaves the app).</summary>
    private List<Dragged>? dragged;
    private string? draggedFrom;

    /// <summary>
    /// A ticked row takes everything ticked with it; any other row goes on its
    /// own. Notes and "show more" don't drag.
    /// </summary>
    private void OnRowsDragStarting(object sender, DragItemsStartingEventArgs e)
    {
        if (active is null || e.Items.FirstOrDefault() is not Row { Node: { } node } row)
        {
            e.Cancel = true;
            return;
        }
        dragged = row.IsChecked
            ? selection.Values.Select(n => new Dragged(n,
                siteOf.TryGetValue(n.Key, out var site) ? site : SpNode.SiteFromUrl(n.SiteUrl ?? n.Url),
                parentOf.GetValueOrDefault(n.Key))).ToList()
            : [new Dragged(node, SiteOf(row), row.Parent?.Node is { Kind: not NodeKind.SiteCollection } parent ? parent : null)];
        draggedFrom = active.Id;
        e.Data.SetText(string.Join(Environment.NewLine, dragged.Select(d => d.Node.Url)));
        e.Data.RequestedOperation = DataPackageOperation.Copy;
    }

    private void OnRowsDragCompleted(ListViewBase sender, DragItemsCompletedEventArgs args)
    {
        dragged = null;
        draggedFrom = null;
        foreach (var item in Nav.MenuItems.OfType<NavigationViewItem>())
        {
            MarkDropTarget(item, false);
        }
    }

    /// <summary>The site collection a row is in: the row at the top of its branch.</summary>
    private static SpNode SiteOf(Row row)
    {
        var root = row;
        while (root.Parent is not null)
        {
            root = root.Parent;
        }
        return root.Node is { Kind: NodeKind.SiteCollection } site ? site : SpNode.SiteFromUrl(row.Node!.SiteUrl ?? row.Node.Url);
    }

    /// <summary>Takes rows from the table, for a saved selection of the same connection. True when it takes them.</summary>
    private bool OnSavedDragOver(string file, DragEventArgs e)
    {
        if (dragged is not { Count: > 0 } items || savedStore.Load(file) is not { } s)
        {
            e.AcceptedOperation = DataPackageOperation.None;
            return false;
        }
        e.DragUIOverride.IsGlyphVisible = false;
        if (s.ConnectionId != draggedFrom)
        {
            e.AcceptedOperation = DataPackageOperation.None;
            e.DragUIOverride.Caption = "Only for a selection of the same connection";
            return false;
        }
        var have = s.Items.Select(n => n.Key).ToHashSet();
        var add = items.Where(d => !have.Contains(d.Node.Key)).ToList();
        e.AcceptedOperation = add.Count == 0 ? DataPackageOperation.None : DataPackageOperation.Copy;
        e.DragUIOverride.Caption = add.Count == 0
            ? $"Already in “{s.Name}”"
            : $"Add {Summary(add.Select(d => d.Node))} to “{s.Name}”";
        return add.Count > 0;
    }

    /// <summary>
    /// The saved selection a drop would go to: a row darker than the one
    /// that's open, in the text's colour (dark on light, light on dark). The
    /// item paints its own background when it isn't the open one.
    /// </summary>
    private static void MarkDropTarget(NavigationViewItem item, bool on)
    {
        if (!on)
        {
            item.ClearValue(Control.BackgroundProperty);
        }
        else if (item.Foreground is SolidColorBrush text && item.Background is not SolidColorBrush { Opacity: DropTargetOpacity })
        {
            item.Background = new SolidColorBrush(text.Color) { Opacity = DropTargetOpacity };
        }
    }

    private const double DropTargetOpacity = 0.16;

    private void OnSavedDrop(string file)
    {
        if (dragged is not { Count: > 0 } items || draggedFrom is null)
        {
            return;
        }
        SavedSelection? updated;
        List<SpNode> added;
        int already;
        try
        {
            if (savedStore.Load(file) is not { } s || s.ConnectionId != draggedFrom)
            {
                return;
            }
            var have = s.Items.Select(n => n.Key).ToHashSet();
            var nodes = items.Select(d => d.Node).DistinctBy(n => n.Key).ToList();
            added = nodes.Where(n => !have.Contains(n.Key)).ToList();
            already = nodes.Count - added.Count;
            updated = savedStore.Add(file, items.Select(d => d.Node), items.Select(d => d.Site), items.Select(d => d.Parent).OfType<SpNode>());
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            ShowError(new InvalidOperationException($"Couldn't add to the saved selection: {ex.Message}"));
            return;
        }
        if (updated is null)
        {
            return;
        }
        var into = updated;
        ShowToast(
            added.Count > 0 ? $"Added to “{into.Name}”" : $"Already in “{into.Name}”",
            string.Join(" · ", new[]
            {
                added.Count > 0 ? Summary(added) : null,
                already > 0 ? $"{already:N0} already in it" : null,
                $"{Plural(into.Items.Count, "item", "items")} in all",
            }.OfType<string>()),
            shownSaved?.File == file ? null : ("Open", () => Nav.SelectedItem = NavItem(SavedTag + file)),
            success: added.Count > 0);
        // the selection on show takes them in too: ticked, as the rest of it
        if (shownSaved?.File == file)
        {
            shownSaved = updated;
            foreach (var d in items)
            {
                selection[d.Node.Key] = d.Node;
                siteOf[d.Node.Key] = d.Site;
                if (d.Parent is { } parent)
                {
                    parentOf[d.Node.Key] = parent;
                }
                foreach (var r in rows.Where(r => r.Node?.Key == d.Node.Key))
                {
                    r.SetChecked(true);
                }
            }
            UpdatePane();
        }
        // after the drop is done with the item it landed on (its badge shows the new count)
        DispatcherQueue.TryEnqueue(RefreshSavedSelections);
    }

    // ---- short confirmations ----

    private Action? toastAction;
    private Microsoft.UI.Dispatching.DispatcherQueueTimer? toastTimer;

    /// <summary>A short confirmation at the bottom of the window, gone after a few seconds (or when closed).</summary>
    private void ShowToast(string title, string message, (string Label, Action Run)? action = null, bool success = true)
    {
        Toast.Title = title;
        Toast.Message = message;
        Toast.Severity = success ? InfoBarSeverity.Success : InfoBarSeverity.Informational;
        toastAction = action?.Run;
        ToastAction.Content = action?.Label;
        ToastAction.Visibility = action is null ? Visibility.Collapsed : Visibility.Visible;
        if (Toast.Shadow is null)
        {
            // lifted off the page, so it reads as a message rather than part of it
            Toast.Shadow = new ThemeShadow();
            Toast.Translation = new System.Numerics.Vector3(0, 0, 32);
        }
        Toast.IsOpen = true;
        toastTimer?.Stop();
        toastTimer = DispatcherQueue.CreateTimer();
        toastTimer.Interval = TimeSpan.FromSeconds(6);
        toastTimer.IsRepeating = false;
        toastTimer.Tick += (_, _) => Toast.IsOpen = false;
        toastTimer.Start();
    }

    private void OnToastAction(object sender, RoutedEventArgs e)
    {
        Toast.IsOpen = false;
        toastAction?.Invoke();
    }

    private void OnShowAll(object sender, RoutedEventArgs e)
    {
        if (identity is not null)
        {
            _ = LoadSitesAsync();
        }
    }
}
