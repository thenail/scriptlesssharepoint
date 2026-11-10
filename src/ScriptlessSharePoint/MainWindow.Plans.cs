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
    private const string PlanTag = "plan:";

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

    // ---- in the navigation pane ----

    /// <summary>The plans below the saved selections, for connections that still exist.</summary>
    private void RefreshPlans()
    {
        var was = navigating;
        navigating = true;
        try
        {
            foreach (var old in Nav.MenuItems.OfType<NavigationViewItem>().Where(i => i.Tag is string t && t.StartsWith(PlanTag, StringComparison.Ordinal)).ToList())
            {
                Nav.MenuItems.Remove(old);
            }
            var connections = workspace.Load().Connections.ToDictionary(c => c.Id, c => c.Name);
            var first = true;
            foreach (var plan in planStore.List().Where(p => connections.ContainsKey(p.ConnectionId)))
            {
                var item = new NavigationViewItem
                {
                    Content = plan.Name,
                    Tag = PlanTag + plan.Id,
                    Icon = new FontIcon { Glyph = "" },
                    Margin = first ? new Thickness(0, 16, 0, 0) : new Thickness(0),
                    InfoBadge = new InfoBadge { Value = plan.Steps.Count },
                };
                ToolTipService.SetToolTip(item, $"Plan: {Plural(plan.Steps.Count, "step", "steps")}{(plan.Prepared is null ? "" : ", a run prepared")}\n" +
                                                $"{connections[plan.ConnectionId]} · made by {plan.CreatedBy ?? "you"} {When(plan.Created)}");
                Nav.MenuItems.Add(item);
                first = false;
            }
            if (shownPlan is { } shown && NavItem(PlanTag + shown.Id) is { } picked)
            {
                Nav.SelectedItem = picked;
            }
        }
        finally
        {
            navigating = was;
        }
    }

    /// <summary>Follows Plans\ in the data folder: plans an AI app makes or changes show up straight away.</summary>
    private void WatchPlans()
    {
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
    /// Opens a plan on what's ticked. With nothing ticked it can start on, its
    /// prepared run (or, without one, its test selection) is ticked instead.
    /// </summary>
    private async Task ShowPlanSelectionAsync(Plan plan)
    {
        if (shownPlan?.Id != plan.Id)
        {
            ResetPlanValues();
        }
        shownPlan = plan;
        if (!PlanCheck.StartsOnAny(plan, selection.Values.Select(n => n.Kind)))
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
        SelectNavQuietly(PlanTag + plan.Id);
        ShowPage(TenantPage);
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
        SelectNavQuietly(PlanTag + plan.Id);
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
        SelectNavQuietly(PlanTag + plan.Id);
    }

    /// <summary>The plan in the pane: its steps, what it runs on, and the values it asks for.</summary>
    private void ShowPlan(Plan plan)
    {
        OverviewView.Visibility = Visibility.Collapsed;
        FormView.Visibility = Visibility.Collapsed;
        ResultsView.Visibility = Visibility.Collapsed;
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
        if (plan.HasTestSelection)
        {
            Link($"Tick its test selection ({(plan.Preset is { } p ? (SitePresets.Find(p)?.Title ?? ContentPresets.Find(p)?.Title ?? p).ToLowerInvariant() : Summary(plan.Items))})",
                async () => { await TickTestSelectionAsync(plan); ShowPlan(plan); });
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
            SelectNavQuietly(PlanTag + plan.Id);
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
        if (items.Count == 0 && !firstOnTenant)
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
        var remove = new Button { Style = (Style)Application.Current.Resources["ToolButton"], Width = 28, Height = 28, Content = new FontIcon { Glyph = "", FontSize = 10 } };
        ToolTipService.SetToolTip(remove, "Take this step out of the plan");
        remove.Click += (_, _) => SavePlan(plan with { Steps = plan.Steps.Where((_, j) => j != i).ToList() }, null);
        Grid.SetColumn(remove, 2);
        grid.Children.Add(remove);
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
        ShowOverview();
    }

    private async void OnDeletePlan(object sender, RoutedEventArgs e)
    {
        if (shownPlan is not { } plan)
        {
            return;
        }
        if (!await ConfirmAsync($"Delete “{plan.Name}”?", "The plan goes from the list. Nothing in SharePoint changes, and its runs stay in the history.", "Delete", danger: false))
        {
            return;
        }
        planStore.Delete(plan.Id);
        shownPlan = null;
        RefreshPlans();
        ShowOverview();
        if (active is not null)
        {
            SelectNavQuietly(active.Id);
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

    // ---- in the pane's command list ----

    /// <summary>The plans that can start on the picked kind, after the commands.</summary>
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
            button.Click += (_, _) => _ = ShowPlanSelectionAsync(plan);
            CommandList.Children.Add(button);
        }
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
        // {variables} typed in the form that the plan doesn't have yet: added as text ones
        var known = PlanValues.BuiltIn.Concat(plan.Variables.Select(v => v.Name)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var added = step.Args.Values.SelectMany(PlanValues.Used).Where(n => !known.Contains(n)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (added.Count > 0)
        {
            plan = plan with { Variables = [.. plan.Variables, .. added.Select(n => new PlanVariable { Name = n })] };
        }
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
