using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.History;
using ScriptlessSharePoint.Plans;
using ScriptlessSharePoint.Selections;
using ScriptlessSharePoint.SharePoint;

using Windows.Storage.Pickers;

namespace ScriptlessSharePoint.App;

/// <summary>
/// Plans: steps (commands with their inputs) to run in order on whatever is
/// ticked. AI apps make them through the MCP server; here they're listed below
/// the saved selections and in the pane (for what's ticked), opened, changed
/// and run. Inputs can hold {variables}: built-in ones from each item, and the
/// plan's own, given here once for all or per item from a pasted table (or
/// prepared by an AI app). A finder step ticks what it finds for the steps
/// after it; found items take the values of the row they're in. Each step's
/// run goes in the jobs' history, which is how AI apps read the results. Plans\ in
/// the data folder is watched, so a plan an AI app has just saved shows up at once.
/// </summary>
public sealed partial class MainWindow
{
    /// <summary>The plans seen so far, with when their prepared run was made: new ones (or new runs) from AI apps are announced.</summary>
    private Dictionary<string, DateTimeOffset?>? knownPlans;

    private readonly PlanStore planStore = new();

    /// <summary>The plan the pane shows.</summary>
    private Plan? shownPlan;

    /// <summary>A plan to open once its connection is signed in.</summary>
    private Plan? pendingPlan;


    // the values the next run of the shown plan uses
    private Dictionary<string, string> planShared = new(StringComparer.OrdinalIgnoreCase);
    private List<PlanRow> planRows = [];
    private string planTableText = "";
    private IReadOnlyList<string> planTableProblems = [];

    private FileSystemWatcher? planWatcher;
    private DispatcherQueueTimer? planChanged;

    // ---- the plans, in the pane ----

    /// <summary>
    /// Whether an open plan came from the pane's list of plans: its back
    /// arrow goes back to the list (else to the commands).
    /// </summary>
    private bool planFromList;

    /// <summary>The plans changed: the pane's list of them, or its commands for what's ticked (plans first).</summary>
    private void RefreshPlans()
    {
        if (PlansListView.Visibility == Visibility.Visible)
        {
            BuildPlansList();
        }
        if (OverviewView.Visibility == Visibility.Visible)
        {
            BuildCommandList();
        }
    }

    /// <summary>
    /// Plans, in the left pane: every plan, in the pane beside the table. The
    /// left pane keeps what it has picked (the selection the plans run on),
    /// and the table stays as it is.
    /// </summary>
    private void ShowPlansList()
    {
        if (active is null)
        {
            return;
        }
        ShowPage(TenantPage);
        // from another page (Jobs, Settings): back to what the table shows, picked
        SelectNavForShown();
        HideReport();
        ShowPane(true);
        editingStep = null;
        planFromList = false;
        OverviewView.Visibility = Visibility.Collapsed;
        FormView.Visibility = Visibility.Collapsed;
        ResultsView.Visibility = Visibility.Collapsed;
        PlanView.Visibility = Visibility.Collapsed;
        PlansListView.Visibility = Visibility.Visible;
        BuildPlansList();
    }

    /// <summary>This connection's plans, then other connections': what each does, what it starts on, who made it, how it last went.</summary>
    private void BuildPlansList()
    {
        PlanCards.Children.Clear();
        var connections = workspace.Load().Connections;
        NewPlanButton.IsEnabled = connections.Count > 0;
        var plans = planStore.List().Where(p => connections.Any(c => c.Id == p.ConnectionId)).ToList();
        if (plans.Count == 0)
        {
            PlanCards.Children.Add(new TextBlock
            {
                Text = "No plans yet. Make one with New plan, or with Add to plan beside a command's Run. AI apps make them through the MCP server (Settings shows how to connect one).",
                Style = (Style)Application.Current.Resources["CaptionSecondary"],
                Margin = new Thickness(4, 4, 4, 0),
            });
            return;
        }
        var runs = historyStore.List().Where(r => r.Plan is not null).ToList();
        foreach (var group in plans.GroupBy(p => p.ConnectionId).OrderBy(g => g.Key == active?.Id ? 0 : 1))
        {
            if (connections.Count > 1)
            {
                PlanCards.Children.Add(new TextBlock
                {
                    Text = (group.Key == active?.Id ? "" : "Other connections: ") + connections.First(c => c.Id == group.Key).Name,
                    Style = (Style)Application.Current.Resources["PaneGroupHeader"],
                    Margin = new Thickness(4, PlanCards.Children.Count == 0 ? 4 : 16, 0, 2),
                });
            }
            foreach (var plan in group)
            {
                PlanCards.Children.Add(PlanCard(plan, runs.Where(r => r.Plan == plan.Id).MaxBy(r => r.Started)));
            }
        }
    }

    private FrameworkElement PlanCard(Plan plan, RunRecord? last)
    {
        var caption = (Style)Application.Current.Resources["CaptionSecondary"];
        var text = new StackPanel { Spacing = 2 };
        text.Children.Add(new TextBlock { Text = plan.Name, Style = (Style)Application.Current.Resources["BodyStrongTextBlockStyle"], TextWrapping = TextWrapping.Wrap });
        if (!string.IsNullOrWhiteSpace(plan.Description))
        {
            text.Children.Add(new TextBlock { Text = plan.Description, TextWrapping = TextWrapping.Wrap });
        }
        var steps = string.Join("  →  ", plan.Steps.Select((st, i) => $"{i + 1}. {CommandCatalog.Find(st.Command)?.Title ?? st.Command}"));
        text.Children.Add(new TextBlock { Text = steps.Length > 0 ? steps : "No steps yet", Style = caption, Margin = new Thickness(0, 4, 0, 0) });
        var starts = PlanCheck.StartKinds(plan);
        text.Children.Add(new TextBlock
        {
            Text = string.Join(" · ", new[]
            {
                plan.Steps.Count == 0 ? null : $"starts on {(starts.Count == 0 ? "the tenant" : string.Join(" or ", starts.Select(k => NodeKinds.Name(k, 2))))}",
                plan.Variables.Count > 0 ? $"asks for {string.Join(", ", plan.Variables.Select(v => v.Shown))}" : null,
                $"made by {plan.CreatedBy ?? "you"} {When(plan.Created)}",
                plan.Prepared is { } prepared ? $"a run prepared by {prepared.By ?? "you"}" : null,
                last is null ? "not run yet" : $"last run {When(last.Started)}: {RunSummary(last)}",
            }.OfType<string>()),
            Style = caption,
        });
        if (PlanCheck.Problems(plan) is { Count: > 0 } problems)
        {
            text.Children.Add(new TextBlock
            {
                Text = problems.Count == 1 ? problems[0] : $"{problems[0]} (and {problems.Count - 1:N0} more)",
                Style = caption,
                Foreground = (Microsoft.UI.Xaml.Media.Brush)Application.Current.Resources["SystemFillColorCautionBrush"],
            });
        }

        var buttons = new Grid { Margin = new Thickness(0, 8, 0, 0) };
        buttons.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        buttons.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        var open = new Button { Content = "Open", MinWidth = 80 };
        ToolTipService.SetToolTip(open, plan.ConnectionId == active?.Id
            ? "Open it here, on what's ticked (or, with nothing ticked, its prepared run or test selection)"
            : "Open it on its own connection");
        open.Click += (_, _) => OpenFromList(plan.Id);
        buttons.Children.Add(open);
        var more = new Button { Style = (Style)Application.Current.Resources["ToolButton"], Content = new FontIcon { Glyph = "", FontSize = 14 }, Flyout = PlanMenu(plan.Id) };
        ToolTipService.SetToolTip(more, "Rename, duplicate or delete");
        Microsoft.UI.Xaml.Automation.AutomationProperties.SetName(more, $"More for {plan.Name}");
        Grid.SetColumn(more, 1);
        buttons.Children.Add(more);
        text.Children.Add(buttons);

        var card = new Border
        {
            Style = (Style)Application.Current.Resources["Card"],
            CornerRadius = (CornerRadius)Application.Current.Resources["ControlCornerRadius"],
            Padding = new Thickness(14, 12, 8, 10),
            Child = text,
            ContextFlyout = PlanMenu(plan.Id),
        };
        card.DoubleTapped += (_, _) => OpenFromList(plan.Id);
        return card;
    }

    /// <summary>A plan opened from the list: its back arrow comes back here.</summary>
    private void OpenFromList(string id)
    {
        _ = OpenPlanFromListAsync();

        async Task OpenPlanFromListAsync()
        {
            await OpenPlanAsync(id);
            planFromList = PlanView.Visibility == Visibility.Visible;
        }
    }

    /// <summary>A plan with no steps, for the active connection (or one picked), opened to add its first step.</summary>
    private async void OnNewPlan(object sender, RoutedEventArgs e)
    {
        var connections = workspace.Load().Connections;
        if (connections.Count == 0)
        {
            return;
        }
        ComboBox? pick = null;
        if (connections.Count > 1)
        {
            pick = new ComboBox { Header = "Connection", HorizontalAlignment = HorizontalAlignment.Stretch, Margin = new Thickness(0, 12, 0, 0) };
            foreach (var c in connections)
            {
                pick.Items.Add(new ComboBoxItem { Content = c.Name, Tag = c.Id });
            }
            pick.SelectedIndex = Math.Max(0, connections.ToList().FindIndex(c => c.Id == active?.Id));
        }
        var name = await AskNameAsync("New plan", $"Plan {DateTime.Now:yyyy-MM-dd HH:mm}", "Create", pick);
        if (name is null)
        {
            return;
        }
        var connectionId = (pick?.SelectedItem as ComboBoxItem)?.Tag as string ?? active?.Id ?? connections[0].Id;
        Plan plan;
        try
        {
            plan = planStore.Save(new Plan { Name = name, ConnectionId = connectionId, CreatedBy = "you" });
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            ShowError(ex);
            return;
        }
        knownPlans?.TryAdd(plan.Id, null);
        OpenFromList(plan.Id);
    }

    /// <summary>Follows Plans\ in the data folder: plans an AI app makes or changes show up straight away.</summary>
    private void WatchPlans()
    {
        knownPlans = planStore.List().ToDictionary(p => p.Id, p => p.Prepared?.At);
        try
        {
            Directory.CreateDirectory(planStore.Folder);
            planWatcher = new FileSystemWatcher(planStore.Folder, "*.json")
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,
                EnableRaisingEvents = true,
            };
            void Changed(object _, FileSystemEventArgs __) => ui.TryEnqueue(() =>
            {
                // a burst of events for one save: act once it's over
                planChanged ??= ui.CreateTimer();
                planChanged.Interval = TimeSpan.FromMilliseconds(300);
                planChanged.IsRepeating = false;
                planChanged.Tick -= OnPlansChanged;
                planChanged.Tick += OnPlansChanged;
                planChanged.Stop();
                planChanged.Start();
            });
            planWatcher.Created += Changed;
            planWatcher.Changed += Changed;
            planWatcher.Deleted += Changed;
            planWatcher.Renamed += (s, e) => Changed(s, e);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException)
        {
            Diagnostics.Log.Write("plans", $"can't watch {planStore.Folder}: {e.Message}");
        }
    }

    private void OnPlansChanged(DispatcherQueueTimer sender, object args)
    {
        AnnounceNewPlans();
        RefreshPlans();
        if (shownPlan is { } shown)
        {
            if (planStore.Load(shown.Id) is { } fresh)
            {
                shownPlan = fresh;
                if (PlanView.Visibility == Visibility.Visible)
                {
                    ShowPlan(fresh);
                }
            }
            else
            {
                shownPlan = null;
                if (PlanView.Visibility == Visibility.Visible)
                {
                    ShowOverview();
                }
            }
        }
        if (OverviewView.Visibility == Visibility.Visible)
        {
            BuildCommandList();
        }
    }

    /// <summary>A plan an AI app has just made, or a run it has just prepared: said in a message, with Open.</summary>
    private void AnnounceNewPlans()
    {
        var plans = planStore.List();
        if (knownPlans is { } known)
        {
            var connections = workspace.Load().Connections.Select(c => c.Id).ToHashSet();
            foreach (var plan in plans.Where(p => connections.Contains(p.ConnectionId)))
            {
                var isNew = !known.TryGetValue(plan.Id, out var preparedAt);
                var by = isNew ? plan.CreatedBy : plan.Prepared?.By;
                if (by is null || by == "you" || (!isNew && plan.Prepared?.At == preparedAt))
                {
                    continue;
                }
                var id = plan.Id;
                ShowToast(isNew ? $"{by} made a plan" : $"{by} prepared a run",
                    isNew
                        ? $"“{plan.Name}”: {Plural(plan.Steps.Count, "step", "steps")}. Open it to look it over and run it."
                        : $"Of “{plan.Name}”: open it to look the run over and start it.",
                    ("Open", () => _ = OpenPlanAsync(id)), success: false);
            }
        }
        knownPlans = plans.ToDictionary(p => p.Id, p => p.Prepared?.At);
    }

    // ---- opening one ----

    private async Task OpenPlanAsync(string id)
    {
        if (planStore.Load(id) is not { } plan)
        {
            RefreshPlans();
            return;
        }
        if (workspace.Load().Find(plan.ConnectionId) is null)
        {
            ShowError(new InvalidOperationException($"The plan \"{plan.Name}\" is for a connection that's gone."));
            return;
        }
        // for another connection: that connection, in the left pane (from the Plans or Jobs page,
        // what the table shows stays picked: ShowPlanSelectionAsync picks it)
        if (active?.Id != plan.ConnectionId)
        {
            SelectNavQuietly(plan.ConnectionId);
        }
        if (active?.Id != plan.ConnectionId || identity is null)
        {
            // signs in to its connection, then shows it
            pendingPlan = plan;
            await ShowConnectionAsync(plan.ConnectionId, force: active?.Id == plan.ConnectionId);
            return;
        }
        await ShowPlanSelectionAsync(plan);
    }

    /// <summary>
    /// Opens a plan on what's ticked, which stays picked in the left pane. With
    /// nothing ticked, its prepared run (or, without one, its test selection)
    /// is ticked instead.
    /// </summary>
    private async Task ShowPlanSelectionAsync(Plan plan)
    {
        if (shownPlan?.Id != plan.Id)
        {
            ResetPlanValues();
        }
        shownPlan = plan;
        // with nothing ticked, its prepared run or test selection; what's picked stays as it is (when it
        // doesn't fit, the plan says so, and "Run it on…" offers the rest)
        if (selection.Count == 0 && !PlanCheck.StartsOnAny(plan, []))
        {
            if (plan.Prepared is { } prepared)
            {
                await UsePreparedAsync(plan, prepared);
            }
            else if (plan.HasTestSelection)
            {
                await TickTestSelectionAsync(plan);
            }
        }
        shownPlan = plan;
        ShowPage(TenantPage);
        // the selection it runs on stays picked in the left pane (its test selection or prepared run: the connection)
        SelectNavForShown();
        ShowPane(true);
        ShowPlan(plan);
    }

    private void ResetPlanValues()
    {
        planShared = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        planRows = [];
        planTableText = "";
        planTableProblems = [];
    }

    /// <summary>Ticks a prepared run's items and takes its values.</summary>
    private async Task UsePreparedAsync(Plan plan, PreparedRun run)
    {
        if (run.Preset is { } preset)
        {
            await ShowPresetAsync(preset, refresh: false);
        }
        else if (active is not null)
        {
            ShowSaved(new SavedSelection { Name = $"{plan.Name}: prepared run", ConnectionId = plan.ConnectionId, Saved = run.At, Items = run.Items, Sites = run.Items });
        }
        planShared = new Dictionary<string, string>(run.Values, StringComparer.OrdinalIgnoreCase);
        planRows = run.Rows.ToList();
        planTableText = planRows.Count > 0 ? PlanValues.TableText(planRows, plan.Variables.Select(v => v.Name)) : "";
        planTableProblems = [];
        shownPlan = plan;
    }

    private async Task TickTestSelectionAsync(Plan plan)
    {
        if (plan.Preset is { } preset)
        {
            await ShowPresetAsync(preset, refresh: false);
        }
        else if (plan.Items.Count > 0 && active is not null)
        {
            ShowSaved(new SavedSelection
            {
                Name = $"{plan.Name}: test selection",
                ConnectionId = plan.ConnectionId,
                Saved = plan.Changed,
                Items = plan.Items,
                Sites = plan.Sites,
                Parents = plan.Parents,
            });
        }
        shownPlan = plan;
    }

    /// <summary>The plan in the pane: its steps, what it runs on, and the values it asks for.</summary>
    private void ShowPlan(Plan plan)
    {
        editingStep = null;
        OverviewView.Visibility = Visibility.Collapsed;
        FormView.Visibility = Visibility.Collapsed;
        ResultsView.Visibility = Visibility.Collapsed;
        PlansListView.Visibility = Visibility.Collapsed;
        PlanView.Visibility = Visibility.Visible;

        PlanTitle.Text = plan.Name;
        PlanDescription.Text = plan.Description ?? "";
        PlanDescription.Visibility = string.IsNullOrWhiteSpace(plan.Description) ? Visibility.Collapsed : Visibility.Visible;
        PlanMeta.Text = $"Made by {plan.CreatedBy ?? "you"} {When(plan.Created)}{(plan.Changed > plan.Created.AddMinutes(1) ? $" · changed {When(plan.Changed)}" : "")}" +
                        $" · starts on {(string.Join(" or ", PlanCheck.StartKinds(plan).Select(k => NodeKinds.Name(k, 2))) is { Length: > 0 } starts ? starts : "the tenant")}";

        var problems = PlanCheck.Problems(plan);
        PlanProblems.IsOpen = problems.Count > 0;
        PlanProblems.Message = string.Join("\n", problems);

        BuildPlanSelectionActions(plan);
        BuildPlanInputs(plan);

        var runs = historyStore.List().Where(r => r.Plan == plan.Id).ToList();
        PlanSteps.Children.Clear();
        for (var i = 0; i < plan.Steps.Count; i++)
        {
            PlanSteps.Children.Add(StepCard(plan, i, runs.Where(r => r.Step == i + 1).MaxBy(r => r.Started)));
        }
        PlanSteps.Children.Add(AddStepButton(plan));
        RefreshPlanCheck();
    }

    /// <summary>Links to tick the plan's test selection or prepared run, and to keep what's ticked.</summary>
    private void BuildPlanSelectionActions(Plan plan)
    {
        PlanSelectionActions.Children.Clear();
        void Link(string text, Action act)
        {
            var link = new HyperlinkButton { Content = text, Padding = new Thickness(0, 2, 0, 2) };
            link.Click += (_, _) => act();
            PlanSelectionActions.Children.Add(link);
        }
        if (plan.Prepared is { } prepared)
        {
            var what = prepared.Preset is { } p
                ? (SitePresets.Find(p)?.Title ?? ContentPresets.Find(p)?.Title ?? p).ToLowerInvariant()
                : Summary(prepared.Items);
            Link($"Use the run {prepared.By ?? "you"} prepared: {what}{(prepared.Rows.Count > 0 ? $", values for {prepared.Rows.Count:N0}" : "")}",
                async () => { await UsePreparedAsync(plan, prepared); ShowPlan(plan); });
        }
        if (RunOnButton(plan) is { } runOn)
        {
            PlanSelectionActions.Children.Add(runOn);
        }
        if (selection.Count > 0)
        {
            Link("Keep what's ticked as its test selection", () => SavePlan(plan with
            {
                Preset = shownPreset is { } preset && shownSaved?.File.Length == 0 ? preset : null,
                Items = shownPreset is not null ? [] : selection.Values.ToList(),
                Sites = shownPreset is not null ? [] : selection.Values.Select(n => siteOf.TryGetValue(n.Key, out var s) ? s : SpNode.SiteFromUrl(n.SiteUrl ?? n.Url)).ToList(),
                Parents = shownPreset is not null ? [] : selection.Values.Select(n => parentOf.GetValueOrDefault(n.Key)).OfType<SpNode>().ToList(),
            }, "Kept as its test selection."));
            if (plan.Variables.Count > 0)
            {
                Link("Keep this run for later (what's ticked, and the values)", () => SavePlan(plan with
                {
                    Prepared = new PreparedRun
                    {
                        Preset = shownPreset,
                        Items = shownPreset is not null ? [] : selection.Values.ToList(),
                        Values = new Dictionary<string, string>(planShared, StringComparer.OrdinalIgnoreCase),
                        Rows = planRows.ToList(),
                        By = "you",
                        At = DateTimeOffset.Now,
                    },
                }, "Kept: it's offered when the plan opens with nothing ticked."));
            }
        }
    }

    /// <summary>
    /// "Run it on…": ticks something for the plan without leaving it: its test
    /// selection, a saved selection of its connection, or a predefined one, of
    /// what the plan starts on.
    /// </summary>
    private DropDownButton? RunOnButton(Plan plan)
    {
        var starts = PlanCheck.StartKinds(plan);
        var menu = new MenuFlyout { Placement = Microsoft.UI.Xaml.Controls.Primitives.FlyoutPlacementMode.BottomEdgeAlignedLeft };
        void Item(IList<MenuFlyoutItemBase> into, string text, string glyph, Func<Task> tick)
        {
            var item = new MenuFlyoutItem { Text = text, Icon = new FontIcon { Glyph = glyph } };
            item.Click += async (_, _) =>
            {
                await tick();
                // the plan stays on show, on what's ticked now
                shownPlan = plan;
                ShowPane(true);
                ShowPlan(plan);
            };
            into.Add(item);
        }
        static string PresetTitle(string id) => SitePresets.Find(id)?.Title ?? ContentPresets.Find(id)?.Title ?? id;

        if (plan.HasTestSelection)
        {
            Item(menu.Items, $"Its test selection: {(plan.Preset is { } p ? PresetTitle(p) : Summary(plan.Items))}", "", () => TickTestSelectionAsync(plan));
        }
        var saved = active is null ? [] : savedStore.List()
            .Where(s => s.ConnectionId == active.Id && s.Items.Any(n => starts.Contains(n.Kind)))
            .ToList();
        if (saved.Count > 0)
        {
            var sub = new MenuFlyoutSubItem { Text = "A saved selection", Icon = new FontIcon { Glyph = "" } };
            foreach (var s in saved)
            {
                Item(sub.Items, s.Preset is null ? $"{s.Name} ({Summary(s.Items)})" : $"{s.Name} (picked again)", s.Preset is null ? "\uE728" : "\uE895", () =>
                {
                    if (s.Preset is { } rule)
                    {
                        return ShowPresetAsync(rule, refresh: false, s);
                    }
                    ShowSaved(s);
                    return Task.CompletedTask;
                });
            }
            menu.Items.Add(sub);
        }
        var sites = starts.Contains(NodeKind.SiteCollection) ? SitePresets.All.ToList() : [];
        var content = starts.Contains(NodeKind.File) || starts.Contains(NodeKind.Item) ? ContentPresets.All.ToList() : [];
        if (sites.Count + content.Count > 0)
        {
            var sub = new MenuFlyoutSubItem { Text = "A predefined selection", Icon = new FontIcon { Glyph = "" } };
            foreach (var preset in sites)
            {
                Item(sub.Items, preset.OneDrive ? $"OneDrives: {preset.Title}" : preset.Title, preset.Glyph, () => ShowPresetAsync(preset.Id, refresh: false));
            }
            if (sites.Count > 0 && content.Count > 0)
            {
                sub.Items.Add(new MenuFlyoutSeparator());
            }
            foreach (var preset in content)
            {
                Item(sub.Items, preset.Title, preset.Glyph, () => ShowPresetAsync(preset.Id, refresh: false));
            }
            menu.Items.Add(sub);
        }
        if (menu.Items.Count == 0)
        {
            return null;
        }
        var button = new DropDownButton { Content = "Run it on…", Flyout = menu, Margin = new Thickness(0, 4, 0, 4) };
        ToolTipService.SetToolTip(button, "Tick a selection for it, and stay in the plan");
        return button;
    }

    /// <summary>A field per variable (a value for all), and a table for a value per item.</summary>
    private void BuildPlanInputs(Plan plan)
    {
        PlanInputs.Children.Clear();
        var used = plan.Steps.SelectMany(s => s.Args.Values.SelectMany(PlanValues.Used)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var variables = plan.Variables.Where(v => used.Contains(v.Name)).ToList();
        if (variables.Count == 0)
        {
            return;
        }
        PlanInputs.Children.Add(new TextBlock { Text = "Values", Style = (Style)Application.Current.Resources["BodyStrongTextBlockStyle"], Margin = new Thickness(0, 6, 0, 0) });
        foreach (var v in variables)
        {
            var box = new TextBox
            {
                Header = $"{v.Shown}  ·  {{{v.Name}}}",
                Text = planShared.GetValueOrDefault(v.Name, ""),
                PlaceholderText = v.Kind switch
                {
                    ParameterKind.Person => "alex@contoso.com",
                    ParameterKind.Number => "A number",
                    ParameterKind.Url => "https://…",
                    _ => "",
                },
                Description = Hint(v.Hint ?? "For every item. A table below can give each its own."),
                AcceptsReturn = v.Kind == ParameterKind.LongText,
            };
            var name = v.Name;
            box.TextChanged += (_, _) =>
            {
                planShared[name] = box.Text;
                RefreshPlanCheck();
            };
            PlanInputs.Children.Add(box);
        }

        var names = variables.Select(v => v.Name).ToList();
        var table = new TextBox
        {
            Text = planTableText,
            AcceptsReturn = true,
            TextWrapping = TextWrapping.NoWrap,
            FontFamily = new FontFamily("Cascadia Mono, Consolas"),
            FontSize = 12,
            Height = 150,
            PlaceholderText = $"URL\t{string.Join("\t", names)}\nhttps://contoso.sharepoint.com/sites/falcon\t…",
        };
        ScrollViewer.SetHorizontalScrollBarVisibility(table, ScrollBarVisibility.Auto);
        table.TextChanged += (_, _) =>
        {
            planTableText = table.Text;
            ReadTable(plan);
            RefreshPlanCheck();
        };
        var load = new Button { Content = "Load a CSV file…" };
        load.Click += async (_, _) =>
        {
            var picker = new FileOpenPicker();
            WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);
            picker.FileTypeFilter.Add(".csv");
            picker.FileTypeFilter.Add(".txt");
            if (await picker.PickSingleFileAsync() is { } file)
            {
                table.Text = await File.ReadAllTextAsync(file.Path);
            }
        };
        var tick = new Button { Content = "Tick the table's sites" };
        ToolTipService.SetToolTip(tick, "Run it on the table's rows: their site collections");
        tick.Click += (_, _) =>
        {
            if (planRows.Count == 0 || active is null)
            {
                return;
            }
            var sites = planRows.Select(r => SpNode.SiteFromUrl(r.Url)).DistinctBy(n => n.Key).ToList();
            ShowSaved(new SavedSelection { Name = $"{plan.Name}: the table's sites", ConnectionId = plan.ConnectionId, Saved = DateTimeOffset.Now, Items = sites, Sites = sites });
            shownPlan = plan;
            ShowPlan(plan);
        };
        var content = new StackPanel { Spacing = 8 };
        content.Children.Add(new TextBlock
        {
            Text = "Paste from Excel with the header row, or load a CSV: the address first, then a column per variable. A row is for that site and everything in it; its values win over the ones above.",
            Style = (Style)Application.Current.Resources["CaptionSecondary"],
            TextWrapping = TextWrapping.Wrap,
        });
        content.Children.Add(table);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        buttons.Children.Add(load);
        buttons.Children.Add(tick);
        content.Children.Add(buttons);
        PlanInputs.Children.Add(new Expander
        {
            Header = planRows.Count > 0 ? $"A value per item: {Plural(planRows.Count, "row", "rows")}" : "A value per item (a table)",
            Content = content,
            IsExpanded = planRows.Count > 0 || planTableText.Length > 0,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            HorizontalContentAlignment = HorizontalAlignment.Stretch,
        });
    }

    /// <summary>The pasted table as rows, and what's wrong with it (columns the plan has no variable for, too).</summary>
    private void ReadTable(Plan plan)
    {
        var (rows, columns, problems) = PlanValues.ParseTable(planTableText);
        var unknown = columns.Where(c => plan.Variables.All(v => !string.Equals(v.Name, c, StringComparison.OrdinalIgnoreCase))).ToList();
        planRows = rows.ToList();
        planTableProblems = unknown.Count > 0
            ? [.. problems, $"The table has columns the plan has no variable for: {string.Join(", ", unknown)}. Its variables: {string.Join(", ", plan.Variables.Select(v => v.Name))}."]
            : problems;
    }

    /// <summary>Whether the shown plan can run on what's ticked with these values; says what's missing.</summary>
    private void RefreshPlanCheck()
    {
        if (shownPlan is not { } plan || PlanView.Visibility != Visibility.Visible)
        {
            return;
        }
        var items = selection.Values.ToList();
        var problems = new List<string>(planTableProblems);
        var starts = PlanCheck.StartKinds(plan);
        var firstOnTenant = plan.Steps.Count > 0 && CommandCatalog.Find(plan.Steps[0].Command)?.AppliesTo(NodeKind.Tenant) == true;
        if (plan.Steps.Count == 0)
        {
            PlanRunsOn.Text = "Add its first step: that decides what it runs on.";
            problems.Add("Add a step.");
        }
        else if (items.Count == 0 && !firstOnTenant)
        {
            PlanRunsOn.Text = $"Nothing ticked: tick the {string.Join(" or ", starts.Select(k => NodeKinds.Name(k, 2)))} to run it on" +
                              (plan.Prepared is not null || plan.HasTestSelection ? ", or use one of these:" : ".");
            problems.Add("Tick what to run it on.");
        }
        else if (items.Count > 0 && !PlanCheck.StartsOnAny(plan, items.Select(n => n.Kind)))
        {
            PlanRunsOn.Text = $"What's ticked ({Summary(items)}) isn't what it starts on: {string.Join(" or ", starts.Select(k => NodeKinds.Name(k, 2)))}.";
            problems.Add("Tick what it starts on.");
        }
        else
        {
            PlanRunsOn.Text = items.Count == 0 ? "The tenant." : $"What's ticked: {Summary(items)}. Untick what it shouldn't run on.";
            var on = items.Count == 0 ? Targets(NodeKind.Tenant) : items;
            var (runProblems, notes) = PlanCheck.RunProblems(plan, on, planShared, planRows, DateTimeOffset.Now);
            problems.AddRange(runProblems);
            if (problems.Count == 0 && notes.Count > 0)
            {
                PlanRunCheck.Severity = InfoBarSeverity.Informational;
                PlanRunCheck.Title = "Note";
                PlanRunCheck.Message = string.Join("\n", notes);
                PlanRunCheck.IsOpen = true;
            }
            else
            {
                PlanRunCheck.IsOpen = false;
            }
        }
        if (problems.Count > 0)
        {
            PlanRunCheck.Severity = InfoBarSeverity.Warning;
            PlanRunCheck.Title = "Before it can run";
            PlanRunCheck.Message = string.Join("\n", problems.Distinct());
            PlanRunCheck.IsOpen = true;
        }
        RunPlanButton.IsEnabled = problems.Count == 0 && PlanCheck.Problems(plan).Count == 0;
    }

    private FrameworkElement StepCard(Plan plan, int i, RunRecord? last)
    {
        var step = plan.Steps[i];
        var command = CommandCatalog.Find(step.Command);
        var grid = new Grid { ColumnSpacing = 10 };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.Children.Add(new Border
        {
            Width = 22,
            Height = 22,
            CornerRadius = new CornerRadius(11),
            Background = (Brush)Application.Current.Resources["AccentFillColorDefaultBrush"],
            VerticalAlignment = VerticalAlignment.Top,
            Child = new TextBlock
            {
                Text = (i + 1).ToString(System.Globalization.CultureInfo.CurrentCulture),
                FontSize = 12,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = (Brush)Application.Current.Resources["TextOnAccentFillColorPrimaryBrush"],
            },
        });
        var text = new StackPanel { Spacing = 2 };
        text.Children.Add(new TextBlock { Text = command?.Title ?? step.Command, TextWrapping = TextWrapping.Wrap });
        var caption = (Style)Application.Current.Resources["CaptionSecondary"];
        if (command is not null)
        {
            var inputs = InputsText(command, PlanCheck.ArgsOf(command, step));
            if (inputs.Length > 0)
            {
                text.Children.Add(new TextBlock { Text = inputs, Style = caption, TextWrapping = TextWrapping.Wrap });
            }
            if (command.After == After.Tick)
            {
                text.Children.Add(new TextBlock { Text = "Ticks what it finds: the steps after it run on that.", Style = caption, TextWrapping = TextWrapping.Wrap });
            }
            if (command.Destructive)
            {
                text.Children.Add(new TextBlock
                {
                    Text = "Changes or removes things in a way that can't always be undone.",
                    Style = caption,
                    Foreground = (Brush)Application.Current.Resources["SystemFillColorCautionBrush"],
                    TextWrapping = TextWrapping.Wrap,
                });
            }
        }
        if (last is not null)
        {
            text.Children.Add(new TextBlock
            {
                Text = $"Last run {When(last.Started)}: {RunSummary(last)}",
                Style = (Style)Application.Current.Resources["CaptionTextBlockStyle"],
                Foreground = (Brush)Application.Current.Resources["TextFillColorTertiaryBrush"],
                TextWrapping = TextWrapping.Wrap,
            });
        }
        Grid.SetColumn(text, 1);
        grid.Children.Add(text);
        var tools = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Top };
        var change = new Button { Style = (Style)Application.Current.Resources["ToolButton"], Width = 28, Height = 28, Content = new FontIcon { Glyph = "", FontSize = 12 } };
        ToolTipService.SetToolTip(change, "Change its inputs");
        Microsoft.UI.Xaml.Automation.AutomationProperties.SetName(change, $"Change step {i + 1}");
        change.IsEnabled = command is not null;
        change.Click += (_, _) => EditStep(plan, i);
        tools.Children.Add(change);
        var more = new Button { Style = (Style)Application.Current.Resources["ToolButton"], Width = 28, Height = 28, Content = new FontIcon { Glyph = "", FontSize = 12 } };
        ToolTipService.SetToolTip(more, "Move or remove");
        Microsoft.UI.Xaml.Automation.AutomationProperties.SetName(more, $"Move or remove step {i + 1}");
        var menu = new MenuFlyout { Placement = Microsoft.UI.Xaml.Controls.Primitives.FlyoutPlacementMode.BottomEdgeAlignedRight };
        void Item(string text, string glyph, bool enabled, Action act)
        {
            var item = new MenuFlyoutItem { Text = text, Icon = new FontIcon { Glyph = glyph }, IsEnabled = enabled };
            item.Click += (_, _) => act();
            menu.Items.Add(item);
        }
        Item("Move up", "", i > 0, () => SavePlan(plan with { Steps = Swap(plan.Steps, i, i - 1) }, null));
        Item("Move down", "", i < plan.Steps.Count - 1, () => SavePlan(plan with { Steps = Swap(plan.Steps, i, i + 1) }, null));
        menu.Items.Add(new MenuFlyoutSeparator());
        Item("Remove the step", "", true, () => SavePlan(plan with { Steps = plan.Steps.Where((_, j) => j != i).ToList() }, $"Step {i + 1} removed."));
        more.Flyout = menu;
        tools.Children.Add(more);
        Grid.SetColumn(tools, 2);
        grid.Children.Add(tools);
        return new Border
        {
            Style = (Style)Application.Current.Resources["Card"],
            CornerRadius = new CornerRadius(6),
            Padding = new Thickness(10, 8, 6, 8),
            Child = grid,
        };
    }

    /// <summary>Saves a change to the shown plan and shows it again.</summary>
    private void SavePlan(Plan plan, string? done)
    {
        try
        {
            shownPlan = planStore.Save(plan);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            ShowError(ex);
            return;
        }
        ShowPlan(shownPlan);
        RefreshPlans();
        if (done is not null)
        {
            ShowToast(shownPlan.Name, done);
        }
    }

    private void OnClosePlan(object sender, RoutedEventArgs e)
    {
        shownPlan = null;
        // opened from the list of plans: back to the list
        if (planFromList)
        {
            ShowPlansList();
            return;
        }
        ShowOverview();
    }

    private async void OnDeletePlan(object sender, RoutedEventArgs e)
    {
        if (shownPlan is { } plan)
        {
            await DeletePlanAsync(plan.Id);
        }
    }

    private async Task DeletePlanAsync(string id)
    {
        if (planStore.Load(id) is not { } plan ||
            !await ConfirmAsync($"Delete “{plan.Name}”?", "The plan goes from the list. Nothing in SharePoint changes, and its runs stay in the history.", "Delete", danger: false))
        {
            return;
        }
        planStore.Delete(plan.Id);
        RefreshPlans();
        if (shownPlan?.Id == id)
        {
            shownPlan = null;
            ShowOverview();
            if (active is not null)
            {
                SelectNavQuietly(active.Id);
            }
        }
    }

    // ---- running ----

    private async void OnRunPlan(object sender, RoutedEventArgs e)
    {
        if (shownPlan is not { } plan || PlanCheck.Problems(plan).Count > 0 || CommandCatalog.Find(plan.Steps[0].Command) is not { } first)
        {
            return;
        }
        var shared = new Dictionary<string, string>(planShared.Where(x => !string.IsNullOrWhiteSpace(x.Value)).ToDictionary(), StringComparer.OrdinalIgnoreCase);
        var rows = planRows.ToList();
        var steps = string.Join("\n", plan.Steps.Select((s, i) =>
        {
            var c = CommandCatalog.Find(s.Command)!;
            var inputs = InputsText(c, PlanCheck.ArgsOf(c, s)).ReplaceLineEndings("; ");
            return $"{i + 1}. {c.Title}{(inputs.Length > 0 ? $" ({inputs})" : "")}";
        }));
        var values = string.Join("; ", plan.Variables.Select(v => rows.Any(r => r.Values.ContainsKey(v.Name))
            ? $"{v.Shown}: from the table ({Plural(rows.Count, "row", "rows")}){(shared.TryGetValue(v.Name, out var s) ? $", else {s}" : "")}"
            : shared.TryGetValue(v.Name, out var all) ? $"{v.Shown}: {all}" : null).OfType<string>());
        var on = selection.Count == 0 ? "the tenant" : Summary(selection.Values);
        if (!await ConfirmAsync($"Run “{plan.Name}”?", $"{plan.Description}\n\n{steps}{(values.Length > 0 ? $"\n\nValues: {values}" : "")}\n\nStarting on {on}.".Trim(), "Run plan",
                danger: plan.Steps.Any(s => CommandCatalog.Find(s.Command)?.Destructive == true)))
        {
            return;
        }

        // a job: it runs in the background, kept as it goes (one record per step)
        var ticked = selection.Values.ToList();
        if (StartJob(first, ticked, PlanCheck.ArgsOf(first, plan.Steps[0]), settings.ConfirmChanges, plan, shared, rows) is { } job)
        {
            WatchJob(job);
        }
    }

    // ---- changing a plan in the app ----

    /// <summary>A plan's menu in the navigation pane.</summary>
    private MenuFlyout PlanMenu(string id)
    {
        var menu = new MenuFlyout();
        void Item(string text, string glyph, Func<Task> act)
        {
            var item = new MenuFlyoutItem { Text = text, Icon = new FontIcon { Glyph = glyph } };
            item.Click += async (_, _) => await act();
            menu.Items.Add(item);
        }
        Item("Rename…", "", () => RenamePlanAsync(id));
        Item("Duplicate", "", () =>
        {
            DuplicatePlan(id);
            return Task.CompletedTask;
        });
        menu.Items.Add(new MenuFlyoutSeparator());
        Item("Delete", "", () => DeletePlanAsync(id));
        return menu;
    }

    private async Task RenamePlanAsync(string id)
    {
        if (planStore.Load(id) is not { } plan || await AskNameAsync("Rename the plan", plan.Name, "Rename") is not { } name || name == plan.Name)
        {
            return;
        }
        try
        {
            var saved = planStore.Save(plan with { Name = name });
            if (shownPlan?.Id == id)
            {
                shownPlan = saved;
                if (PlanView.Visibility == Visibility.Visible)
                {
                    ShowPlan(saved);
                }
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            ShowError(ex);
        }
        RefreshPlans();
    }

    /// <summary>A copy of a plan, to change without touching the one it's from (a prepared run stays with the original).</summary>
    private void DuplicatePlan(string id)
    {
        if (planStore.Load(id) is not { } plan)
        {
            return;
        }
        try
        {
            var copy = planStore.Save(plan with { File = "", Name = $"{plan.Name} (copy)", CreatedBy = "you", Created = default, Prepared = null });
            RefreshPlans();
            ShowToast($"Made “{copy.Name}”", "A copy to change: the original stays as it is.", ("Open", () => _ = OpenPlanAsync(copy.Id)));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            ShowError(ex);
        }
    }

    private static List<PlanStep> Swap(IReadOnlyList<PlanStep> steps, int a, int b)
    {
        var list = steps.ToList();
        (list[a], list[b]) = (list[b], list[a]);
        return list;
    }

    /// <summary>What's ticked when a plan reaches a step: what it starts on, then what each finder before it ticks.</summary>
    private static IReadOnlySet<NodeKind> KindsBefore(Plan plan, int index)
    {
        var kinds = PlanCheck.StartKinds(plan);
        for (var i = 0; i < index && i < plan.Steps.Count; i++)
        {
            if (CommandCatalog.Find(plan.Steps[i].Command) is { } c)
            {
                kinds = PlanCheck.KindsAfter(c, PlanCheck.ArgsOf(c, plan.Steps[i]), kinds);
            }
        }
        return kinds;
    }

    /// <summary>The kind a command runs on at a step (any but the tenant, when it can).</summary>
    private static NodeKind StepKind(Command c, IReadOnlySet<NodeKind> kinds)
    {
        var on = PlanCheck.RunsOn(c, kinds);
        return on.Count == 0 ? c.Targets[0] : on.FirstOrDefault(k => k != NodeKind.Tenant, on[0]);
    }

    /// <summary>A step's inputs in the form, to change and save (not run).</summary>
    private void EditStep(Plan plan, int index)
    {
        var step = plan.Steps[index];
        if (CommandCatalog.Find(step.Command) is not { } c)
        {
            return;
        }
        // the first step decides what the plan starts on: it can be on any of its command's kinds
        var kinds = index == 0 && plan.StartsOn.Count == 0 ? c.Targets.ToHashSet() : KindsBefore(plan, index);
        editingStep = (plan, index, StepKind(c, kinds));
        OpenForm(c, PlanCheck.ArgsOf(c, step));
    }

    /// <summary>
    /// "Add a step": a box to find one by what it's called, what it does or
    /// the PowerShell it stands in for, over the commands that can run where
    /// the plan has got to, by group. Enter adds the first that matches.
    /// </summary>
    private FrameworkElement AddStepButton(Plan plan)
    {
        IReadOnlySet<NodeKind> kinds = plan.Steps.Count == 0 && plan.StartsOn.Count == 0
            ? Enum.GetValues<NodeKind>().ToHashSet()
            : KindsBefore(plan, plan.Steps.Count);
        var commands = CommandCatalog.All.Where(c => PlanCheck.RunsOn(c, kinds).Count > 0).ToList();

        var flyout = new Flyout { Placement = Microsoft.UI.Xaml.Controls.Primitives.FlyoutPlacementMode.BottomEdgeAlignedLeft };
        var find = new AutoSuggestBox { PlaceholderText = "Find a step", QueryIcon = new SymbolIcon(Symbol.Find), Margin = new Thickness(0, 0, 0, 8) };
        var list = new StackPanel();
        var scroll = new ScrollViewer { Content = list, MaxHeight = 420, Padding = new Thickness(0, 0, 12, 0) };
        var panel = new StackPanel { Width = 340 };
        panel.Children.Add(find);
        panel.Children.Add(scroll);
        flyout.Content = panel;

        void Add(Command c)
        {
            flyout.Hide();
            editingStep = (plan, plan.Steps.Count, StepKind(c, kinds));
            OpenForm(c);
        }
        List<Command> Matching()
        {
            var text = find.Text?.Trim() ?? "";
            return commands.Where(c => text.Length == 0 ||
                                       c.Title.Contains(text, StringComparison.CurrentCultureIgnoreCase) ||
                                       c.Description.Contains(text, StringComparison.CurrentCultureIgnoreCase) ||
                                       c.PowerShell?.Contains(text, StringComparison.OrdinalIgnoreCase) == true).ToList();
        }
        void Fill()
        {
            list.Children.Clear();
            var shown = Matching();
            if (shown.Count == 0)
            {
                list.Children.Add(new TextBlock
                {
                    Text = $"No step that can run here matches “{find.Text?.Trim()}”.",
                    Style = (Style)Application.Current.Resources["CaptionSecondary"],
                    Margin = new Thickness(4, 4, 4, 4),
                });
                return;
            }
            var critical = (Brush)Application.Current.Resources["SystemFillColorCriticalBrush"];
            foreach (var group in shown.GroupBy(c => c.Group))
            {
                list.Children.Add(new TextBlock { Text = group.Key.ToUpper(System.Globalization.CultureInfo.CurrentCulture), Style = (Style)Application.Current.Resources["PaneGroupHeader"] });
                foreach (var c in group)
                {
                    var row = new Grid { ColumnSpacing = 12 };
                    row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                    row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    var icon = new FontIcon { Glyph = c.Glyph, FontSize = 16, VerticalAlignment = VerticalAlignment.Top, Margin = new Thickness(0, 2, 0, 0) };
                    if (c.Destructive)
                    {
                        icon.Foreground = critical;
                    }
                    row.Children.Add(icon);
                    var text = new StackPanel();
                    text.Children.Add(new TextBlock { Text = c.Title, TextWrapping = TextWrapping.Wrap });
                    if (c.PowerShell is { } ps)
                    {
                        text.Children.Add(new TextBlock { Text = ps, Style = (Style)Application.Current.Resources["CaptionSecondary"], TextTrimming = TextTrimming.CharacterEllipsis, TextWrapping = TextWrapping.NoWrap });
                    }
                    Grid.SetColumn(text, 1);
                    row.Children.Add(text);
                    var button = new Button { Style = (Style)Application.Current.Resources["PaneCommand"], Content = row };
                    ToolTipService.SetToolTip(button, c.Description);
                    Microsoft.UI.Xaml.Automation.AutomationProperties.SetName(button, c.Title);
                    button.Click += (_, _) => Add(c);
                    list.Children.Add(button);
                }
            }
        }
        find.TextChanged += (_, _) => Fill();
        find.QuerySubmitted += (_, _) =>
        {
            if (Matching().FirstOrDefault() is { } first)
            {
                Add(first);
            }
        };
        flyout.Opening += (_, _) =>
        {
            find.Text = "";
            Fill();
        };
        flyout.Opened += (_, _) => find.Focus(FocusState.Programmatic);

        var content = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        content.Children.Add(new FontIcon { Glyph = "", FontSize = 14 });
        content.Children.Add(new TextBlock { Text = "Add a step" });
        var button = new DropDownButton { Content = content, Flyout = flyout, Margin = new Thickness(0, 4, 0, 0) };
        Microsoft.UI.Xaml.Automation.AutomationProperties.SetName(button, "Add a step");
        ToolTipService.SetToolTip(button, kinds.Count == 0
            ? "Commands for the tenant"
            : $"Commands for what's ticked by then: {string.Join(", ", kinds.Select(k => NodeKinds.Name(k, 2)))}");
        return button;
    }

    /// <summary>Saves the form as the step being edited (or a new one), with any {variables} it brings.</summary>
    private void SaveStep(Command c)
    {
        if (editingStep is not { } edit)
        {
            return;
        }
        var args = FormArgs();
        var problems = c.Check(PlanValues.WithSamples(c, args), [edit.Kind]);
        FormProblems.Text = string.Join("\n", problems);
        FormProblems.Visibility = problems.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        if (problems.Count > 0)
        {
            return;
        }
        var step = new PlanStep { Command = c.Id, Args = args.ToDictionary().Where(a => a.Value is not null).ToDictionary(a => a.Key, a => a.Value) };
        var plan = planStore.Load(edit.Plan.Id) ?? edit.Plan;
        var steps = plan.Steps.ToList();
        var adding = edit.Index >= steps.Count;
        if (adding)
        {
            steps.Add(step);
        }
        else
        {
            steps[edit.Index] = step;
        }
        var (updated, added) = WithVariablesOf(plan with { Steps = steps }, step);
        editingStep = null;
        SavePlan(updated, $"{(adding ? "Added" : "Changed")} step {edit.Index + 1}: {c.Title}" +
                          (added.Count > 0 ? $", with the variables {string.Join(", ", added.Select(n => $"{{{n}}}"))}." : "."));
    }

    /// <summary>The {variables} a step uses that the plan doesn't have yet, added as text ones.</summary>
    private static (Plan Plan, List<string> Added) WithVariablesOf(Plan plan, PlanStep step)
    {
        var known = PlanValues.BuiltIn.Concat(plan.Variables.Select(v => v.Name)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var added = step.Args.Values.SelectMany(PlanValues.Used).Where(n => !known.Contains(n)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        return added.Count == 0 ? (plan, added) : (plan with { Variables = [.. plan.Variables, .. added.Select(n => new PlanVariable { Name = n })] }, added);
    }

    // ---- in the pane's command list ----

    /// <summary>The plans that can start on the picked kind, before the commands.</summary>
    private void AddPlansToCommandList(string filter)
    {
        if (active is null)
        {
            return;
        }
        var plans = planStore.List()
            .Where(p => p.ConnectionId == active.Id && PlanCheck.StartsOnAny(p, [pickedKind]))
            .Where(p => filter.Length == 0 || p.Name.Contains(filter, StringComparison.CurrentCultureIgnoreCase) || p.Description?.Contains(filter, StringComparison.CurrentCultureIgnoreCase) == true)
            .ToList();
        if (plans.Count == 0)
        {
            return;
        }
        CommandList.Children.Add(new TextBlock { Text = "PLANS", Style = (Style)Application.Current.Resources["PaneGroupHeader"] });
        foreach (var plan in plans)
        {
            var grid = new Grid { ColumnSpacing = 12 };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.Children.Add(new FontIcon { Glyph = "", FontSize = 16, VerticalAlignment = VerticalAlignment.Top, Margin = new Thickness(0, 2, 0, 0) });
            var text = new StackPanel();
            text.Children.Add(new TextBlock { Text = plan.Name, TextWrapping = TextWrapping.Wrap });
            text.Children.Add(new TextBlock
            {
                Text = $"{Plural(plan.Steps.Count, "step", "steps")}{(plan.Variables.Count > 0 ? $" · asks for {string.Join(", ", plan.Variables.Select(v => v.Shown))}" : "")}",
                Style = (Style)Application.Current.Resources["CaptionSecondary"],
                TextTrimming = TextTrimming.CharacterEllipsis,
                TextWrapping = TextWrapping.NoWrap,
            });
            Grid.SetColumn(text, 1);
            grid.Children.Add(text);
            var button = new Button { Style = (Style)Application.Current.Resources["PaneCommand"], Content = grid };
            ToolTipService.SetToolTip(button, plan.Description ?? plan.Name);
            Microsoft.UI.Xaml.Automation.AutomationProperties.SetName(button, $"Plan: {plan.Name}");
            button.Click += (_, _) =>
            {
                planFromList = false;
                _ = ShowPlanSelectionAsync(plan);
            };
            CommandList.Children.Add(button);
        }
        var all = new HyperlinkButton { Content = "All plans…", Margin = new Thickness(12, 0, 0, 4), Padding = new Thickness(0, 2, 0, 2) };
        all.Click += (_, _) => ShowPlansList();
        CommandList.Children.Add(all);
    }

    // ---- adding a step from a command's form ----

    private void OnAddToPlanOpening(object sender, object e)
    {
        AddToPlanMenu.Items.Clear();
        var create = new MenuFlyoutItem { Text = "New plan…", Icon = new FontIcon { Glyph = "" } };
        create.Click += async (_, _) => await AddToPlanAsync(null);
        AddToPlanMenu.Items.Add(create);
        var plans = active is null ? [] : planStore.List().Where(p => p.ConnectionId == active.Id).Take(15).ToList();
        if (plans.Count > 0)
        {
            AddToPlanMenu.Items.Add(new MenuFlyoutSeparator());
        }
        foreach (var plan in plans)
        {
            var item = new MenuFlyoutItem { Text = plan.Name, Icon = new FontIcon { Glyph = "" } };
            item.Click += async (_, _) => await AddToPlanAsync(plan.Id);
            AddToPlanMenu.Items.Add(item);
        }
    }

    /// <summary>The form's command and inputs as a step, added to a plan (or a new one, with what's ticked as its test selection).</summary>
    private async Task AddToPlanAsync(string? planId)
    {
        if (command is not { } c || active is null)
        {
            return;
        }
        var args = FormArgs();
        var problems = c.Check(PlanValues.WithSamples(c, args), [pickedKind]);
        FormProblems.Text = string.Join("\n", problems);
        FormProblems.Visibility = problems.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        if (problems.Count > 0)
        {
            return;
        }
        var step = new PlanStep { Command = c.Id, Args = args.ToDictionary().Where(a => a.Value is not null).ToDictionary(a => a.Key, a => a.Value) };
        Plan plan;
        if (planId is null)
        {
            var name = await AskNameAsync("New plan", $"{c.Title}, {DateTime.Now:yyyy-MM-dd}", "Create");
            if (name is null)
            {
                return;
            }
            var items = pickedKind == NodeKind.Tenant ? [] : selection.Values.Where(n => n.Kind == pickedKind).ToList();
            plan = new Plan
            {
                Name = name,
                ConnectionId = active.Id,
                CreatedBy = "you",
                StartsOn = pickedKind == NodeKind.Tenant ? [] : [pickedKind],
                Items = items,
                Sites = items.Select(n => siteOf.TryGetValue(n.Key, out var site) ? site : SpNode.SiteFromUrl(n.SiteUrl ?? n.Url)).ToList(),
                Parents = items.Select(n => parentOf.GetValueOrDefault(n.Key)).OfType<SpNode>().ToList(),
                Steps = [step],
            };
        }
        else if (planStore.Load(planId) is { } existing)
        {
            plan = existing with { Steps = [.. existing.Steps, step] };
        }
        else
        {
            return;
        }
        (plan, var added) = WithVariablesOf(plan, step);
        try
        {
            plan = planStore.Save(plan);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            ShowError(ex);
            return;
        }
        RefreshPlans();
        var saved = plan;
        var issues = PlanCheck.Problems(saved);
        ShowToast(planId is null ? $"Made the plan “{saved.Name}”" : $"Added step {saved.Steps.Count} to “{saved.Name}”",
            issues.Count > 0 ? issues[0] : $"{c.Title}{(added.Count > 0 ? $", with the variables {string.Join(", ", added.Select(n => $"{{{n}}}"))}" : "")}. Open the plan to run it.",
            ("Open", () => _ = OpenPlanAsync(saved.Id)), success: issues.Count == 0);
    }
}
