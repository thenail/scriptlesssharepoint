using System.Diagnostics;
using System.Runtime.InteropServices;

using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;

using ScriptlessSharePoint.Auth;
using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.Connections;
using ScriptlessSharePoint.Diagnostics;
using ScriptlessSharePoint.Settings;
using ScriptlessSharePoint.SharePoint;

using Windows.ApplicationModel.DataTransfer;
using Windows.Graphics;
using Windows.Storage.Pickers;
using VirtualKey = Windows.System.VirtualKey;
using VirtualKeyModifiers = Windows.System.VirtualKeyModifiers;

using WinRT.Interop;

namespace ScriptlessSharePoint.App;

/// <summary>
/// The window: connections down the left, a connection's site collections in
/// the table (MainWindow.Table.cs), and the pane that runs commands on what's
/// ticked (MainWindow.Pane.cs). Settings has the AI apps and About.
/// </summary>
public sealed partial class MainWindow : Window, IRowOwner, ICommandHost
{
    private readonly IntPtr hwnd;
    private readonly DispatcherQueue ui;
    private readonly Workspace workspace;
    private ConnectionProfile? active;
    private Identity? identity;
    private bool navigating;
    private int busy;

    public MainWindow()
    {
        InitializeComponent();
        Log.Init("app");
        Log.Write("app", $"{AppInfo.Name} {AppInfo.Version} starting from {AppContext.BaseDirectory}");
        hwnd = WindowNative.GetWindowHandle(this);
        ui = DispatcherQueue.GetForCurrentThread();
        workspace = new Workspace(new ConnectionStore(), () => hwnd);
        workspace.AccountChanged += p => ui.TryEnqueue(() => OnAccountChanged(p));

        SetUpChrome();
        Table.ItemsSource = rows;
        Table.ContainerContentChanging += OnRowShown;
        Table.SizeChanged += (_, e) => Columns.Current.Fit(e.NewSize.Width);
        Root.KeyboardAcceleratorPlacementMode = KeyboardAcceleratorPlacementMode.Hidden;
        AddAccelerator(VirtualKey.F5, VirtualKeyModifiers.None, () => OnRefresh(this, new RoutedEventArgs()));
        AddAccelerator(VirtualKey.F, VirtualKeyModifiers.Control, () => SearchBox.Focus(FocusState.Keyboard));
        AddAccelerator(VirtualKey.S, VirtualKeyModifiers.Control, () => _ = SaveSelectionAsync());
        AddAccelerator(VirtualKey.S, VirtualKeyModifiers.Control | VirtualKeyModifiers.Shift, () => _ = SaveSelectionAsync(asNew: true));
        // Ctrl+Z ticks again what was ticked before, except in a text box, where it's the text's
        var undo = new KeyboardAccelerator { Key = VirtualKey.Z, Modifiers = VirtualKeyModifiers.Control };
        undo.Invoked += (_, e) =>
        {
            if (!TypingInText() && TenantPage.Visibility == Visibility.Visible)
            {
                e.Handled = true;
                UndoTicks();
            }
        };
        Root.KeyboardAccelerators.Add(undo);
        Root.Loaded += (_, _) => Start();
        // jobs running as the app closes are kept as they are: they come back paused
        Closed += (_, _) => SaveNow();
    }

    private void SetUpChrome()
    {
        // the title lives in the title bar, as in the Settings app
        ExtendsContentIntoTitleBar = true;
        SetTitleBar(AppTitleBar);
        AppWindow.TitleBar.PreferredHeightOption = TitleBarHeightOption.Tall;
        var icon = Path.Combine(AppContext.BaseDirectory, "Assets", "app.ico");
        if (File.Exists(icon))
        {
            AppWindow.SetIcon(icon);
        }

        var scale = GetDpiForWindow(hwnd) / 96.0;
        var area = DisplayArea.GetFromWindowId(AppWindow.Id, DisplayAreaFallback.Primary).WorkArea;
        var width = (int)Math.Min(1480 * scale, area.Width * 0.92);
        var height = (int)Math.Min(920 * scale, area.Height * 0.92);
        AppWindow.MoveAndResize(new RectInt32(area.X + (area.Width - width) / 2, area.Y + (area.Height - height) / 2, width, height));
        if (AppWindow.Presenter is OverlappedPresenter presenter)
        {
            presenter.PreferredMinimumWidth = (int)(980 * scale);
            presenter.PreferredMinimumHeight = (int)(600 * scale);
        }
    }

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hwnd);

    private void AddAccelerator(VirtualKey key, VirtualKeyModifiers modifiers, Action action)
    {
        var accelerator = new KeyboardAccelerator { Key = key, Modifiers = modifiers };
        accelerator.Invoked += (_, e) =>
        {
            e.Handled = true;
            action();
        };
        Root.KeyboardAccelerators.Add(accelerator);
    }

    private void Start()
    {
        // the app is in English: so is Settings, whatever Windows' language
        if (Nav.SettingsItem is NavigationViewItem settingsItem)
        {
            settingsItem.Content = "Settings";
        }
        VersionText.Text = $"{AppInfo.Name} {AppInfo.Version}";
        DataFolderText.Text = AppInfo.DataDir;
        LoadSettings();
        BuildAiSnippets();
        if (App.StartSite is { } site)
        {
            SearchBox.Text = site;
        }
        ReloadNav(workspace.Load().Active?.Id);
        UpdatePane();
        WatchPlans();
        RestoreJobs();
        if (App.StartPlan is { } plan)
        {
            ui.TryEnqueue(() => _ = OpenPlanAsync(plan));
        }
    }

    // ---- connections down the left ----

    /// <summary>The connections as navigation items, selecting one (which shows it).</summary>
    private void ReloadNav(string? select)
    {
        var data = workspace.Load();
        navigating = true;
        try
        {
            // the predefined selections stay open or closed as they were
            KeepPresetsOpen();
            Nav.MenuItems.Clear();
            if (data.Connections.Count > 0)
            {
                Nav.MenuItems.Add(NavHeader("Connections", "header:connections"));
            }
            foreach (var c in data.Connections)
            {
                var item = new NavigationViewItem
                {
                    Content = c.Name,
                    Tag = c.Id,
                    Icon = new FontIcon { Glyph = c.Kind switch { ConnectionKind.Certificate => "", ConnectionKind.AccessToken => "", _ => "" } },
                };
                ToolTipService.SetToolTip(item, $"{c.TenantUrl}\n{c.KindLabel}");
                item.ContextFlyout = ConnectionMenu(c);
                Nav.MenuItems.Add(item);
            }
            if (data.Connections.Count > 0)
            {
                AddPresets();
            }
            RefreshSavedSelections();
            RefreshPlans();
        }
        finally
        {
            navigating = false;
        }

        if (data.Connections.Count == 0)
        {
            active = null;
            identity = null;
            UpdateAccount();
            ShowPage(WelcomePage);
            return;
        }
        var target = Nav.MenuItems.OfType<NavigationViewItem>().FirstOrDefault(i => (string)i.Tag == select)
            ?? Nav.MenuItems.OfType<NavigationViewItem>().First();
        if (ReferenceEquals(Nav.SelectedItem, target))
        {
            _ = ShowConnectionAsync((string)target.Tag, force: true);
        }
        else
        {
            Nav.SelectedItem = target;
        }
    }

    /// <summary>A heading over a part of the navigation pane, tagged so the part can be rebuilt with it.</summary>
    private static NavigationViewItemHeader NavHeader(string text, string tag) => new() { Content = text, Tag = tag };

    private MenuFlyout ConnectionMenu(ConnectionProfile c)
    {
        var menu = new MenuFlyout();
        var edit = new MenuFlyoutItem { Text = "Edit connection…", Icon = new FontIcon { Glyph = "" } };
        edit.Click += (_, _) => _ = EditConnectionAsync(c.Id);
        menu.Items.Add(edit);
        if (c.Kind == ConnectionKind.User)
        {
            var signOut = new MenuFlyoutItem { Text = "Sign out", Icon = new FontIcon { Glyph = "" } };
            signOut.Click += (_, _) => _ = SignOutAsync(c.Id);
            menu.Items.Add(signOut);
        }
        menu.Items.Add(new MenuFlyoutSeparator());
        var delete = new MenuFlyoutItem { Text = "Delete connection", Icon = new FontIcon { Glyph = "" } };
        delete.Click += (_, _) => _ = DeleteConnectionAsync(c);
        menu.Items.Add(delete);
        return menu;
    }

    private void OnNavSelectionChanged(NavigationView sender, NavigationViewSelectionChangedEventArgs args)
    {
        if (navigating)
        {
            return;
        }
        // a plan on show stays in the pane: the left pane picks what's ticked, not what to do with it
        if (args.IsSettingsSelected)
        {
            ShowPage(SettingsPage);
        }
        else if (args.SelectedItem is NavigationViewItem { Tag: HistoryTag })
        {
            ShowHistory();
        }
        else if (args.SelectedItem is NavigationViewItem { Tag: PlansTag })
        {
            ShowPlansPage();
        }
        else if (args.SelectedItem is NavigationViewItem { Tag: string preset } && preset.StartsWith(PresetTag, StringComparison.Ordinal))
        {
            _ = OpenPresetAsync(preset[PresetTag.Length..]);
        }
        else if (args.SelectedItem is NavigationViewItem { Tag: string tag } && tag.StartsWith(SavedTag, StringComparison.Ordinal))
        {
            _ = OpenSavedAsync(tag[SavedTag.Length..]);
        }
        else if (args.SelectedItem is NavigationViewItem { Tag: string id })
        {
            pendingSaved = null;
            _ = ShowConnectionAsync(id);
        }
    }

    private void OnNavItemInvoked(NavigationView sender, NavigationViewItemInvokedEventArgs args)
    {
        if (args.InvokedItemContainer == AddConnectionItem)
        {
            OnAddConnection(sender, new RoutedEventArgs());
        }
        else if (args.InvokedItemContainer is NavigationViewItem { Tag: string id } && id == active?.Id)
        {
            ShowPage(TenantPage);
            // the connection itself, while a saved selection shows: every site collection again
            if (shownSaved is not null && identity is not null)
            {
                _ = LoadSitesAsync();
            }
        }
    }

    private void ShowPage(FrameworkElement page)
    {
        foreach (var p in new FrameworkElement[] { WelcomePage, TenantPage, PlansPage, JobsPage, SettingsPage })
        {
            p.Visibility = p == page ? Visibility.Visible : Visibility.Collapsed;
        }
        SearchBox.IsEnabled = page == TenantPage;
    }

    /// <summary>Shows a connection: who it's signed in as (found silently), then its site collections.</summary>
    private async Task ShowConnectionAsync(string id, bool force = false)
    {
        ShowPage(TenantPage);
        if (!force && active?.Id == id)
        {
            return;
        }
        var v = ++view;
        HideError();
        // a plan of another connection doesn't stay on show
        if (shownPlan is { } plan && plan.ConnectionId != id)
        {
            shownPlan = null;
        }
        try
        {
            workspace.SetActive(id);
            active = workspace.Load().Find(id);
        }
        catch (ConnectionException e)
        {
            ShowError(e);
            return;
        }
        if (active is null)
        {
            return;
        }
        identity = null;
        adminScope = true;
        ResetTable();
        ResetSelection();
        SetCrumbs(null);
        TenantLink.Content = active.TenantUrl.Replace("https://", "");
        TenantDetail.Text = active.KindLabel;
        UpdateAccount();
        SetStatus("Checking sign-in…", "", working: true);
        ShowEmpty("Checking sign-in…", ring: true);

        Identity? who = null;
        try
        {
            who = Demo.On ? Demo.Identity : await workspace.StatusAsync();
        }
        catch (Exception e)
        {
            ShowError(e);
        }
        if (v != view)
        {
            return;
        }
        identity = who;
        UpdateAccount();
        if (identity is null)
        {
            ShowSignedOut();
            return;
        }
        await ShowSitesOrSavedAsync();
    }

    private void UpdateAccount()
    {
        if (identity is null)
        {
            AccountPicture.DisplayName = "";
            AccountName.Text = active is null ? "No connection" : "Not signed in";
            AccountDetail.Text = active?.AccountName ?? active?.Name ?? "";
            return;
        }
        AccountPicture.DisplayName = identity.IsApp ? "App" : identity.DisplayName;
        AccountName.Text = identity.IsApp ? $"App-only: {identity.DisplayName}" : identity.DisplayName;
        AccountDetail.Text = identity.Email ?? identity.TenantTitle;
        SignInButton.Visibility = Visibility.Collapsed;
    }

    private void ShowSignedOut()
    {
        SetStatus("Not signed in", "");
        ShowEmpty($"Sign in to {active?.Name} to see its site collections.");
        SignInButton.Visibility = Visibility.Visible;
    }

    // ---- where you are ----

    /// <summary>The page's title: the connection, then (when there is one) what the table shows instead of every site collection.</summary>
    private void SetCrumbs(string? shown)
    {
        if (active is null)
        {
            return;
        }
        Crumbs.ItemsSource = shown is null ? new[] { active.Name } : new[] { active.Name, shown };
        ToolTipService.SetToolTip(Crumbs, shown is null ? null : $"Pick {active.Name} to see every site collection again");
    }

    /// <summary>The connection in the title: every site collection again (no search).</summary>
    private void OnCrumbClicked(BreadcrumbBar sender, BreadcrumbBarItemClickedEventArgs args)
    {
        if (args.Index != 0 || identity is null || sender.ItemsSource is not string[] { Length: > 1 })
        {
            return;
        }
        typing?.Stop();
        SearchBox.Text = "";
        _ = LoadSitesAsync();
    }

    private async void OnSignIn(object sender, RoutedEventArgs e)
    {
        if (active is null)
        {
            return;
        }
        HideError();
        var v = view;
        SetStatus("Signing in…", "", working: true);
        try
        {
            identity = await workspace.SignInAsync();
        }
        catch (Exception ex)
        {
            ShowError(ex);
            ShowSignedOut();
            return;
        }
        if (v != view)
        {
            return;
        }
        UpdateAccount();
        await ShowSitesOrSavedAsync();
    }

    private async Task SignOutAsync(string id)
    {
        try
        {
            if (active?.Id != id)
            {
                await ShowConnectionAsync(id);
            }
            await workspace.SignOutAsync();
            identity = null;
            ++view;
            ResetTable();
            ResetSelection();
            UpdateAccount();
            ShowSignedOut();
        }
        catch (Exception e)
        {
            ShowError(e);
        }
    }

    private void OnAccountChanged(ConnectionProfile profile)
    {
        if (active?.Id == profile.Id)
        {
            active = workspace.Load().Find(profile.Id) ?? active;
            if (identity is null)
            {
                UpdateAccount();
            }
        }
    }

    private async void OnAddConnection(object sender, RoutedEventArgs e)
    {
        var dialog = new ConnectionDialog(workspace, null, this) { XamlRoot = Root.XamlRoot };
        await dialog.ShowAsync();
        if (dialog.Saved is { } saved)
        {
            ReloadNav(saved.Id);
        }
    }

    private async Task EditConnectionAsync(string id)
    {
        var profile = workspace.Load().Find(id);
        if (profile is null)
        {
            return;
        }
        var dialog = new ConnectionDialog(workspace, profile, this) { XamlRoot = Root.XamlRoot };
        await dialog.ShowAsync();
        if (dialog.Saved is { } saved)
        {
            active = null; // shown again with the saved settings
            ReloadNav(saved.Id);
        }
    }

    private async Task DeleteConnectionAsync(ConnectionProfile c)
    {
        var saved = savedStore.List().Count(s => s.ConnectionId == c.Id);
        if (!await ConfirmAsync($"Delete “{c.Name}”?",
                "The connection and its settings are removed from this PC. AI apps can't use it any more either. The account stays signed in to Windows." +
                (saved == 0 ? "" : $" {(saved == 1 ? "Its saved selection goes" : $"Its {saved} saved selections go")} too."),
                "Delete", danger: true))
        {
            return;
        }
        try
        {
            savedStore.DeleteFor(c.Id);
            var data = workspace.Delete(c.Id);
            if (active?.Id == c.Id)
            {
                active = null;
            }
            ReloadNav(data.ActiveId);
        }
        catch (Exception e)
        {
            ShowError(e);
        }
    }

    private void OnTenantLink(object sender, RoutedEventArgs e)
    {
        if (active is not null)
        {
            OpenUrl(active.TenantUrl);
        }
    }

    // ---- status and errors ----

    private void SetStatus(string title, string text, bool working = false)
    {
        StatusTitle.Text = title;
        StatusTitle.Visibility = title.Length == 0 ? Visibility.Collapsed : Visibility.Visible;
        StatusText.Text = text;
        BusyBar.Visibility = working ? Visibility.Visible : Visibility.Collapsed;
    }

    private void Busy(bool on)
    {
        busy += on ? 1 : -1;
        BusyBar.Visibility = busy > 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private void ShowError(Exception e)
    {
        var message = e is ConnectionException or ArgumentException ? e.Message : SharePointErrors.Describe(e);
        var signIn = e is SignInRequiredException || e.InnerException is SignInRequiredException;
        Log.Write("app", $"error: {e.GetType().Name}: {message}");
        ErrorBar.Message = message;
        ErrorSignIn.Visibility = signIn ? Visibility.Visible : Visibility.Collapsed;
        ErrorBar.IsOpen = true;
    }

    private void HideError() => ErrorBar.IsOpen = false;

    // ---- dialogs and pickers ----

    internal async Task<bool> ConfirmAsync(string title, string body, string action, bool danger)
    {
        var dialog = new ContentDialog
        {
            XamlRoot = Root.XamlRoot,
            Title = title,
            Content = new TextBlock { Text = body, TextWrapping = TextWrapping.Wrap },
            PrimaryButtonText = action,
            CloseButtonText = "Cancel",
            DefaultButton = danger ? ContentDialogButton.Close : ContentDialogButton.Primary,
        };
        return await dialog.ShowAsync() == ContentDialogResult.Primary;
    }

    internal async Task<string?> PickFileAsync(params string[] types)
    {
        var picker = new FileOpenPicker { ViewMode = PickerViewMode.List, SuggestedStartLocation = PickerLocationId.DocumentsLibrary };
        foreach (var t in types.DefaultIfEmpty("*"))
        {
            picker.FileTypeFilter.Add(t);
        }
        InitializeWithWindow.Initialize(picker, hwnd);
        return (await picker.PickSingleFileAsync())?.Path;
    }

    internal async Task<string?> PickFolderAsync()
    {
        var picker = new FolderPicker { SuggestedStartLocation = PickerLocationId.Downloads };
        picker.FileTypeFilter.Add("*");
        InitializeWithWindow.Initialize(picker, hwnd);
        return (await picker.PickSingleFolderAsync())?.Path;
    }

    internal async Task<string?> PickSaveFileAsync(string name, string description, string extension)
    {
        var picker = new FileSavePicker { SuggestedStartLocation = PickerLocationId.DocumentsLibrary, SuggestedFileName = name };
        picker.FileTypeChoices.Add(description, [extension]);
        InitializeWithWindow.Initialize(picker, hwnd);
        return (await picker.PickSaveFileAsync())?.Path;
    }

    // ---- ICommandHost ----

    public void OpenUrl(string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeHttps)
        {
            return;
        }
        try
        {
            Process.Start(new ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true })?.Dispose();
        }
        catch (Exception e) when (e is System.ComponentModel.Win32Exception or InvalidOperationException)
        {
            Log.Write("app", $"couldn't open {uri.GetLeftPart(UriPartial.Authority)}: {e.Message}");
        }
    }

    public void CopyText(string text) => ui.TryEnqueue(() =>
    {
        var data = new DataPackage();
        data.SetText(text);
        Clipboard.SetContent(data);
    });

    // ---- settings ----

    private readonly SettingsStore settingsStore = new();
    private AppSettings settings = new();

    private void LoadSettings()
    {
        settings = settingsStore.Load();
        ConfirmChangesSwitch.IsOn = settings.ConfirmChanges;
    }

    private void OnConfirmChangesToggled(object sender, RoutedEventArgs e)
    {
        if (ConfirmChangesSwitch.IsOn == settings.ConfirmChanges)
        {
            return;
        }
        settings = settings with { ConfirmChanges = ConfirmChangesSwitch.IsOn };
        try
        {
            settingsStore.Save(settings);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            ShowError(ex);
        }
    }

    /// <summary>How AI apps start the MCP server: the execution alias when installed, else the exe next to this one.</summary>
    private static string McpCommand
    {
        get
        {
            var dir = AppContext.BaseDirectory;
            if (dir.Contains(@"\WindowsApps\", StringComparison.OrdinalIgnoreCase))
            {
                return AppInfo.McpAlias;
            }
            var sibling = Path.Combine(dir, AppInfo.McpExe);
            return File.Exists(sibling) ? sibling : AppInfo.McpAlias;
        }
    }

    private void BuildAiSnippets()
    {
        const string name = "scriptless-sharepoint";
        var command = McpCommand;
        var json = System.Text.Json.JsonSerializer.Serialize(command);
        (string Title, string How, string Code)[] snippets =
        [
            ("Claude Desktop", "Settings › Developer › Edit Config, add this to claude_desktop_config.json, then restart Claude.",
                $"{{\n  \"mcpServers\": {{\n    \"{name}\": {{ \"command\": {json} }}\n  }}\n}}"),
            ("Claude Code", "Run this in a terminal.", $"claude mcp add {name} -- {(command.Contains(' ') ? $"\"{command}\"" : command)}"),
            ("VS Code", "Add this to .vscode/mcp.json, or run MCP: Add Server.",
                $"{{\n  \"servers\": {{\n    \"{name}\": {{ \"type\": \"stdio\", \"command\": {json} }}\n  }}\n}}"),
            ("Windows", "In the Microsoft Store version, Windows' own AI agents find the server without setup.", ""),
        ];
        foreach (var (title, how, code) in snippets)
        {
            var text = new StackPanel { Spacing = 4 };
            text.Children.Add(new TextBlock { Text = title });
            text.Children.Add(new TextBlock { Text = how, Style = (Style)Application.Current.Resources["CaptionSecondary"] });
            if (code.Length > 0)
            {
                text.Children.Add(new Border
                {
                    Background = (Microsoft.UI.Xaml.Media.Brush)Application.Current.Resources["SubtleFillColorSecondaryBrush"],
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(10, 8, 10, 8),
                    Margin = new Thickness(0, 6, 0, 0),
                    Child = new TextBlock { Text = code, Style = (Style)Application.Current.Resources["Code"] },
                });
            }
            var grid = new Grid { ColumnSpacing = 16 };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.Children.Add(text);
            if (code.Length > 0)
            {
                var copy = new Button { Content = "Copy", VerticalAlignment = VerticalAlignment.Top };
                copy.Click += (_, _) =>
                {
                    CopyText(code);
                    copy.Content = "Copied";
                };
                Grid.SetColumn(copy, 1);
                grid.Children.Add(copy);
            }
            AiSnippets.Children.Add(new Border
            {
                Style = (Style)Application.Current.Resources["Card"],
                CornerRadius = (CornerRadius)Application.Current.Resources["ControlCornerRadius"],
                Padding = new Thickness(16, 14, 16, 14),
                Child = grid,
            });
        }
    }

    private void OnOpenDataFolder(object sender, RoutedEventArgs e)
    {
        Directory.CreateDirectory(AppInfo.DataDir);
        Process.Start(new ProcessStartInfo("explorer.exe", $"\"{AppInfo.DataDir}\"") { UseShellExecute = true })?.Dispose();
    }
}
