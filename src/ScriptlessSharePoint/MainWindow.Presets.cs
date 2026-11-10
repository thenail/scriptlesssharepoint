using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

using ScriptlessSharePoint.Selections;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.App;

/// <summary>
/// Predefined selections, listed below the connections and above the saved
/// selections, in three parts: site collections and OneDrives picked by a rule
/// over the admin centre's properties (read for all of them, kept a few
/// minutes), and files, pages and lists found by SharePoint search. Picking
/// one shows what it picks, ticked, as a saved selection shows its items
/// (each file under its library and site); F5 runs it again, and Save keeps it.
/// </summary>
public sealed partial class MainWindow
{
    private const string PresetTag = "preset:";
    private const string PresetsTag = "presets";

    /// <summary>The predefined selection on show (its ID), so F5 can run it again.</summary>
    private string? shownPreset;

    /// <summary>One to show once the connection is signed in.</summary>
    private string? pendingPreset;

    /// <summary>Which parts of the section are open, by tag; kept while the pane is rebuilt.</summary>
    private readonly Dictionary<string, bool> presetsOpen = new()
    {
        [PresetsTag + ":sites"] = true,
    };

    /// <summary>Notes which parts are open, before the pane is rebuilt.</summary>
    private void KeepPresetsOpen()
    {
        foreach (var item in NavItems(Nav.MenuItems).Where(i => i.Tag is string t && t.StartsWith(PresetsTag, StringComparison.Ordinal)))
        {
            presetsOpen[(string)item.Tag] = item.IsExpanded;
        }
    }

    /// <summary>The three parts in the navigation pane, after the connections (with a gap) and before the saved selections.</summary>
    private void AddPresets()
    {
        var sites = Part(PresetsTag + ":sites", "Site collections", "", "Picked from the admin centre's site properties (needs a SharePoint administrator)");
        sites.Margin = new Thickness(0, 16, 0, 0);
        var oneDrives = Part(PresetsTag + ":onedrives", "OneDrives", "", "Picked from the admin centre's OneDrive properties (needs a SharePoint administrator)");
        var content = Part(PresetsTag + ":content", "Files and pages", "", "Found by SharePoint search: a few minutes behind, and only what you can open");
        foreach (var preset in SitePresets.All)
        {
            (preset.OneDrive ? oneDrives : sites).MenuItems.Add(Rule(preset.Id, preset.Title, preset.Glyph, preset.Description));
        }
        foreach (var preset in ContentPresets.All)
        {
            content.MenuItems.Add(Rule(preset.Id, preset.Title, preset.Glyph, preset.Description));
        }
        Nav.MenuItems.Add(sites);
        Nav.MenuItems.Add(oneDrives);
        Nav.MenuItems.Add(content);

        NavigationViewItem Part(string tag, string title, string glyph, string tip)
        {
            var item = new NavigationViewItem
            {
                Content = title,
                Tag = tag,
                Icon = new FontIcon { Glyph = glyph },
                SelectsOnInvoked = false,
                IsExpanded = presetsOpen.GetValueOrDefault(tag),
            };
            ToolTipService.SetToolTip(item, tip);
            return item;
        }

        static NavigationViewItem Rule(string id, string title, string glyph, string tip)
        {
            var item = new NavigationViewItem { Content = title, Tag = PresetTag + id, Icon = new FontIcon { Glyph = glyph } };
            ToolTipService.SetToolTip(item, tip);
            return item;
        }
    }

    private async Task OpenPresetAsync(string id)
    {
        if (active is null || (SitePresets.Find(id) is null && ContentPresets.Find(id) is null))
        {
            return;
        }
        if (identity is null)
        {
            // signs in, then shows it
            pendingPreset = id;
            await ShowConnectionAsync(active.Id, force: true);
            return;
        }
        await ShowPresetAsync(id, refresh: false);
    }

    /// <summary>What a predefined selection picks, shown ticked.</summary>
    private async Task ShowPresetAsync(string id, bool refresh)
    {
        var v = ++view;
        shownPreset = id;
        shownSaved = null;
        SelectNavQuietly(PresetTag + id);
        typing?.Stop();
        SearchBox.Text = "";
        query = null;
        ShowPage(TenantPage);
        HideError();
        ShowAllButton.Visibility = Visibility.Visible;
        try
        {
            if (SitePresets.Find(id) is { } sites)
            {
                await ShowSitePresetAsync(sites, refresh, v);
            }
            else if (ContentPresets.Find(id) is { } content)
            {
                await ShowContentPresetAsync(content, v);
            }
        }
        catch (Exception e)
        {
            if (v != view)
            {
                return;
            }
            ShowError(SharePointClient.IsAccessDenied(e) && SitePresets.Find(id) is not null
                ? new InvalidOperationException("This reads every site collection's (or OneDrive's) properties from the admin centre: that needs a SharePoint administrator.")
                : e);
            SetStatus("Couldn't find them", "");
            ShowEmpty("Couldn't find them.");
        }
    }

    private async Task ShowSitePresetAsync(SitePreset preset, bool refresh, int v)
    {
        var what = preset.OneDrive ? "OneDrives" : "site collections";
        ShowEmpty($"Reading the {what} from the admin centre…", ring: true);
        SetStatus(preset.Title, $"Reading the {what}…", working: true);
        IReadOnlyList<SiteFacts> all;
        if (Demo.On)
        {
            all = Demo.Facts();
        }
        else
        {
            var progress = new Progress<int>(n =>
            {
                if (v == view)
                {
                    EmptyText.Text = $"Reading the {what} from the admin centre… {n:N0}";
                }
            });
            all = await workspace.Current().Client.SiteFactsAsync(refresh, progress, oneDrives: preset.OneDrive);
        }
        if (v != view || active is null)
        {
            return;
        }
        var now = DateTimeOffset.UtcNow;
        var sites = all.Where(f => preset.Matches(f, now)).Select(f => SpNode.FromSite(f.Summary)).ToList();
        Show(preset.Id, new SavedSelection { Name = Name(preset.Title, preset.OneDrive), ConnectionId = active.Id, Saved = now, Items = sites, Sites = sites });
        SetStatus(Name(preset.Title, preset.OneDrive), $"{sites.Count:N0} of {all.Count:N0} {what}. {preset.Description}");
        if (sites.Count == 0)
        {
            ShowEmpty($"None of the {all.Count:N0} {what}. {preset.Description}");
        }

        static string Name(string title, bool oneDrive) => oneDrive ? $"OneDrives: {title}" : title;
    }

    private async Task ShowContentPresetAsync(ContentPreset preset, int v)
    {
        ShowEmpty("Searching…", ring: true);
        SetStatus(preset.Title, "Searching…", working: true);
        if (Demo.On)
        {
            ShowEmpty("Search isn't part of the demo: its sites are made up.");
            SetStatus(preset.Title, "");
            return;
        }
        var progress = new Progress<int>(n =>
        {
            if (v == view)
            {
                EmptyText.Text = $"Searching… {n:N0} found";
            }
        });
        var now = DateTimeOffset.UtcNow;
        var hits = await workspace.Current().Client.SearchContentAsync(preset.QueryAt(now), preset.SortBy, preset.Descending, ContentPresets.Limit, progress);
        if (v != view || active is null)
        {
            return;
        }
        var found = hits.Select(ContentPresets.Rows).OfType<FoundHit>().DistinctBy(h => h.Node.Key).ToList();
        var sites = found.Select(h => h.Site).DistinctBy(n => n.Key).ToList();
        Show(preset.Id, new SavedSelection
        {
            Name = preset.Title,
            ConnectionId = active.Id,
            Saved = now,
            Items = found.Select(h => h.Node).ToList(),
            Sites = sites,
            Parents = found.SelectMany(h => new[] { h.Parent, h.Web }).OfType<SpNode>().DistinctBy(n => n.Key).ToList(),
        });
        SetStatus(preset.Title,
            $"{Summary(found.Select(h => h.Node))} in {Plural(sites.Count, "site collection", "site collections")}" +
            $"{(hits.Count >= ContentPresets.Limit ? $" (the first {ContentPresets.Limit:N0})" : "")}. {preset.Description} From search: a few minutes behind, and only what you can open.");
        if (found.Count == 0)
        {
            ShowEmpty($"Search found none. {preset.Description}");
        }
    }

    /// <summary>Shows what a predefined selection picked, and keeps it as the one on show.</summary>
    private void Show(string id, SavedSelection picked)
    {
        ShowSaved(picked);
        shownPreset = id;
        SelectNavQuietly(PresetTag + id);
        ShowAllButton.Visibility = Visibility.Visible;
    }
}
