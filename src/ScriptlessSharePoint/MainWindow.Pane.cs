using System.Globalization;
using System.Text;

using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.Connections;
using ScriptlessSharePoint.History;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.App;

/// <summary>
/// The pane beside the table, as File Size's Selected pane: what's ticked, and
/// the commands for one kind of it (or for the tenant). A command opens its
/// inputs here; Run starts it as a job on every ticked thing of that kind,
/// and the results follow the job (Close leaves it running in the background).
/// </summary>
public sealed partial class MainWindow
{
    /// <summary>The kind the pane acts on; Tenant when nothing's ticked.</summary>
    private NodeKind pickedKind = NodeKind.Tenant;
    private bool fillingKinds;
    private Command? command;
    private readonly List<(CommandParameter Parameter, FrameworkElement Control, Func<string?> Value)> fields = [];

    /// <summary>The job the pane's results follow.</summary>
    private Job? watchedJob;
    private JobBinding? paneBinding;

    private IEnumerable<NodeKind> SelectedKinds => selection.Values.Select(n => n.Kind).Distinct().Order();

    /// <summary>What a command on <paramref name="kind"/> runs on.</summary>
    private List<SpNode> Targets(NodeKind kind) => kind == NodeKind.Tenant
        ? active is null ? [] : [SpNode.Tenant(active.EffectiveAdminUrl, identity?.TenantTitle ?? active.Name)]
        : selection.Values.Where(n => n.Kind == kind).ToList();

    private static string TargetsText(NodeKind kind, int count) => kind == NodeKind.Tenant ? "the tenant" : NodeKinds.Count(kind, count);

    /// <summary>Brings the pane up to date with the ticks.</summary>
    private void UpdatePane()
    {
        var count = selection.Count;
        SelectionBadge.Value = count;
        SelectionBadge.Visibility = count > 0 ? Visibility.Visible : Visibility.Collapsed;
        ClearButton.Visibility = count > 0 ? Visibility.Visible : Visibility.Collapsed;
        SaveSelectionButton.Visibility = count > 0 ? Visibility.Visible : Visibility.Collapsed;
        UpdateSaveButton();
        NothingSelected.Visibility = count == 0 ? Visibility.Visible : Visibility.Collapsed;
        SelectedSummary.Text = count == 0
            ? "Nothing ticked"
            : string.Join(" · ", selection.Values.GroupBy(n => n.Kind).OrderBy(g => g.Key).Select(g => NodeKinds.Count(g.Key, g.Count())));

        SelectedExpander.Visibility = count > 0 ? Visibility.Visible : Visibility.Collapsed;
        SelectedExpanderHeader.Text = count == 1 ? selection.Values.First().Title : $"Show the {count:N0} ticked";
        if (SelectedExpander.IsExpanded || count <= 50)
        {
            FillSelectedList();
        }

        // the kinds to act on: those ticked, then the tenant
        var kinds = SelectedKinds.Append(NodeKind.Tenant).ToList();
        if (!kinds.Contains(pickedKind))
        {
            pickedKind = kinds[0];
        }
        fillingKinds = true;
        KindPicker.Items.Clear();
        foreach (var k in kinds)
        {
            var n = k == NodeKind.Tenant ? 0 : selection.Values.Count(x => x.Kind == k);
            KindPicker.Items.Add(new ComboBoxItem { Content = k == NodeKind.Tenant ? "The tenant" : $"{NodeKinds.Heading(k)} ({n:N0})", Tag = k });
        }
        KindPicker.SelectedIndex = kinds.IndexOf(pickedKind);
        fillingKinds = false;

        if (FormView.Visibility == Visibility.Visible && command is not null)
        {
            if (command.AppliesTo(pickedKind) && Targets(pickedKind).Count > 0)
            {
                FormTargets.Text = $"Runs on {TargetsText(pickedKind, Targets(pickedKind).Count)}";
                ShowFields();
            }
            else
            {
                ShowOverview();
            }
        }
        if (OverviewView.Visibility == Visibility.Visible)
        {
            BuildCommandList();
        }
        // a plan on show: what it runs on is what's ticked
        if (PlanView.Visibility == Visibility.Visible)
        {
            RefreshPlanCheck();
        }
    }

    private void FillSelectedList()
    {
        SelectedList.Items.Clear();
        foreach (var node in selection.Values)
        {
            var grid = new Grid { ColumnSpacing = 8, Padding = new Thickness(8, 0, 0, 0), Height = 32 };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.Children.Add(new FontIcon { Glyph = Glyphs.For(node.Kind), FontSize = 14, VerticalAlignment = VerticalAlignment.Center });
            var title = new TextBlock { Text = node.Title, TextTrimming = TextTrimming.CharacterEllipsis, VerticalAlignment = VerticalAlignment.Center };
            ToolTipService.SetToolTip(title, node.Url);
            Grid.SetColumn(title, 1);
            grid.Children.Add(title);
            var remove = new Button { Style = (Style)Application.Current.Resources["ToolButton"], Width = 28, Height = 28, Content = new FontIcon { Glyph = "", FontSize = 10 } };
            ToolTipService.SetToolTip(remove, "Untick");
            Microsoft.UI.Xaml.Automation.AutomationProperties.SetName(remove, $"Untick {node.Title}");
            var key = node.Key;
            remove.Click += (_, _) =>
            {
                Untick(key);
                UpdatePane();
            };
            Grid.SetColumn(remove, 2);
            grid.Children.Add(remove);
            SelectedList.Items.Add(grid);
        }
    }

    private void OnKindChanged(object sender, SelectionChangedEventArgs e)
    {
        if (fillingKinds || KindPicker.SelectedItem is not ComboBoxItem { Tag: NodeKind kind })
        {
            return;
        }
        pickedKind = kind;
        ShowOverview();
        BuildCommandList();
    }

    private void OnCommandFilterChanged(AutoSuggestBox sender, AutoSuggestBoxTextChangedEventArgs args) => BuildCommandList();

    /// <summary>The commands for the picked kind, in groups, filtered by what's typed.</summary>
    private void BuildCommandList()
    {
        CommandList.Children.Clear();
        var filter = CommandFilter.Text?.Trim() ?? "";
        var commands = CommandCatalog.For(pickedKind)
            .Where(c => filter.Length == 0 ||
                        c.Title.Contains(filter, StringComparison.CurrentCultureIgnoreCase) ||
                        c.Description.Contains(filter, StringComparison.CurrentCultureIgnoreCase) ||
                        c.PowerShell?.Contains(filter, StringComparison.OrdinalIgnoreCase) == true)
            .ToList();
        if (commands.Count == 0)
        {
            CommandList.Children.Add(new TextBlock
            {
                Text = $"No command for {NodeKinds.Name(pickedKind, 2)} matches “{filter}”.",
                Style = (Style)Application.Current.Resources["CaptionSecondary"],
                Margin = new Thickness(12, 8, 12, 0),
            });
            AddPlansToCommandList(filter);
            return;
        }
        var critical = (Brush)Application.Current.Resources["SystemFillColorCriticalBrush"];
        foreach (var group in commands.GroupBy(c => c.Group))
        {
            CommandList.Children.Add(new TextBlock { Text = group.Key.ToUpper(CultureInfo.CurrentCulture), Style = (Style)Application.Current.Resources["PaneGroupHeader"] });
            foreach (var c in group)
            {
                var grid = new Grid { ColumnSpacing = 12 };
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                var icon = new FontIcon { Glyph = c.Glyph, FontSize = 16, VerticalAlignment = VerticalAlignment.Top, Margin = new Thickness(0, 2, 0, 0) };
                if (c.Destructive)
                {
                    icon.Foreground = critical;
                }
                grid.Children.Add(icon);
                var text = new StackPanel();
                text.Children.Add(new TextBlock { Text = c.Title, TextWrapping = TextWrapping.Wrap });
                var caption = c.PowerShell ?? "";
                if (c.NeedsAdmin)
                {
                    caption = caption.Length == 0 ? "Administrator" : $"{caption} · administrator";
                }
                if (caption.Length > 0)
                {
                    text.Children.Add(new TextBlock { Text = caption, Style = (Style)Application.Current.Resources["CaptionSecondary"], TextTrimming = TextTrimming.CharacterEllipsis, TextWrapping = TextWrapping.NoWrap });
                }
                Grid.SetColumn(text, 1);
                grid.Children.Add(text);
                var button = new Button { Style = (Style)Application.Current.Resources["PaneCommand"], Content = grid };
                ToolTipService.SetToolTip(button, c.Description);
                Microsoft.UI.Xaml.Automation.AutomationProperties.SetName(button, c.Title);
                button.Click += (_, _) => OpenForm(c);
                CommandList.Children.Add(button);
            }
        }
        AddPlansToCommandList(filter);
    }

    // ---- a command's inputs ----

    /// <summary>Opens a command's form, with <paramref name="values"/> (from the history) in place of its defaults.</summary>
    private void OpenForm(Command c, CommandArgs? values = null)
    {
        command = c;
        var targets = Targets(pickedKind);
        FormTitle.Text = c.Title;
        FormDescription.Text = c.Description;
        FormPowerShell.Text = c.PowerShell is null ? "" : $"In PowerShell: {c.PowerShell}";
        FormPowerShell.Visibility = c.PowerShell is null ? Visibility.Collapsed : Visibility.Visible;
        FormNote.IsOpen = c.Destructive || c.NeedsAdmin;
        FormNote.Severity = c.Destructive ? InfoBarSeverity.Warning : InfoBarSeverity.Informational;
        FormNote.Message = c.Destructive
            ? "This changes or removes things in a way that can't always be undone. You'll be asked first."
            : "Needs a SharePoint administrator, or an app with Sites.FullControl.All.";
        FormProblems.Visibility = Visibility.Collapsed;
        FormTargets.Text = $"Runs on {TargetsText(pickedKind, targets.Count)}";
        RunButton.Content = c.After == After.Tick ? "Tick" : c.ReadOnly ? "Show" : "Run";

        FormFields.Children.Clear();
        fields.Clear();
        var defaults = c.Defaults();
        foreach (var p in c.Parameters)
        {
            var (control, value) = Field(p, values?[p.Name] ?? defaults[p.Name], ShowFields);
            FormFields.Children.Add(control);
            fields.Add((p, control, value));
        }
        ShowFields();

        OverviewView.Visibility = Visibility.Collapsed;
        ResultsView.Visibility = Visibility.Collapsed;
        PlanView.Visibility = Visibility.Collapsed;
        FormView.Visibility = Visibility.Visible;
        if (FormFields.Children.FirstOrDefault() is Control first)
        {
            first.Focus(FocusState.Programmatic);
        }
    }

    /// <summary>
    /// What the form's inputs say now: those that don't apply (to another
    /// input's value, or to what's ticked) are left out.
    /// </summary>
    private CommandArgs FormArgs()
    {
        var args = new CommandArgs();
        foreach (var (p, _, value) in fields)
        {
            args[p.Name] = value();
        }
        foreach (var (p, _, _) in fields)
        {
            if (!p.IsShown(args, pickedKind))
            {
                args[p.Name] = null;
            }
        }
        return args;
    }

    /// <summary>Shows the inputs that apply, as choices and switches change.</summary>
    private void ShowFields()
    {
        var args = new CommandArgs();
        foreach (var (p, _, value) in fields)
        {
            args[p.Name] = value();
        }
        foreach (var (p, control, _) in fields)
        {
            control.Visibility = p.IsShown(args, pickedKind) ? Visibility.Visible : Visibility.Collapsed;
        }
    }

    /// <summary>A hint under an input, wrapped to the pane's width.</summary>
    private static TextBlock? Hint(string? hint) => hint is null ? null : new TextBlock { Text = hint, TextWrapping = TextWrapping.Wrap };

    /// <summary>The control for one input, and how to read it; <paramref name="changed"/> runs when a choice or switch changes.</summary>
    private (FrameworkElement Control, Func<string?> Value) Field(CommandParameter p, string? value, Action changed)
    {
        switch (p.Kind)
        {
            case ParameterKind.Toggle:
            {
                var toggle = new ToggleSwitch { Header = p.Label, IsOn = value == "true" };
                toggle.Toggled += (_, _) => changed();
                return (WithHint(toggle, p.Hint), () => toggle.IsOn ? "true" : "false");
            }
            case ParameterKind.Choice:
            {
                var box = new ComboBox { Header = p.Label, HorizontalAlignment = HorizontalAlignment.Stretch, Description = Hint(p.Hint) };
                foreach (var choice in p.Choices)
                {
                    box.Items.Add(choice.Label);
                }
                box.SelectedIndex = Math.Max(0, p.Choices.ToList().FindIndex(x => x.Value == value));
                box.SelectionChanged += (_, _) => changed();
                return (box, () => box.SelectedIndex >= 0 ? p.Choices[box.SelectedIndex].Value : null);
            }
            case ParameterKind.Number:
            {
                var number = new NumberBox
                {
                    Header = p.Label,
                    Description = Hint(p.Hint),
                    PlaceholderText = p.Placeholder ?? "",
                    Value = long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : double.NaN,
                    SpinButtonPlacementMode = NumberBoxSpinButtonPlacementMode.Hidden,
                    ValidationMode = NumberBoxValidationMode.Disabled,
                };
                return (number, () => double.IsNaN(number.Value) ? null : ((long)Math.Round(number.Value)).ToString(CultureInfo.InvariantCulture));
            }
            case ParameterKind.LocalFile or ParameterKind.LocalFolder:
            {
                var box = new TextBox { Header = p.Label, Description = Hint(p.Hint), Text = value ?? "", PlaceholderText = p.Kind == ParameterKind.LocalFile ? "Choose a file" : "Choose a folder" };
                var browse = new Button { Content = "Browse…", VerticalAlignment = VerticalAlignment.Bottom };
                browse.Click += async (_, _) =>
                {
                    var path = p.Kind == ParameterKind.LocalFile ? await PickFileAsync() : await PickFolderAsync();
                    if (path is not null)
                    {
                        box.Text = path;
                    }
                };
                var grid = new Grid { ColumnSpacing = 8 };
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                grid.Children.Add(box);
                Grid.SetColumn(browse, 1);
                grid.Children.Add(browse);
                return (grid, () => box.Text);
            }
            default:
            {
                var box = new TextBox
                {
                    Header = p.Label,
                    Description = Hint(p.Hint),
                    PlaceholderText = p.Placeholder ?? "",
                    Text = value ?? "",
                };
                if (p.Kind == ParameterKind.LongText)
                {
                    box.AcceptsReturn = true;
                    box.TextWrapping = TextWrapping.Wrap;
                    box.MinHeight = 72;
                }
                if (p.Kind is ParameterKind.Url or ParameterKind.Person)
                {
                    box.IsSpellCheckEnabled = false;
                }
                return (box, () => box.Text);
            }
        }
    }

    private static FrameworkElement WithHint(FrameworkElement control, string? hint)
    {
        if (hint is null)
        {
            return control;
        }
        var panel = new StackPanel();
        panel.Children.Add(control);
        panel.Children.Add(new TextBlock { Text = hint, Style = (Style)Application.Current.Resources["CaptionSecondary"], TextWrapping = TextWrapping.Wrap });
        return panel;
    }

    private void OnBackToCommands(object sender, RoutedEventArgs e) => ShowOverview();

    private void ShowOverview()
    {
        FormView.Visibility = Visibility.Collapsed;
        ResultsView.Visibility = Visibility.Collapsed;
        PlanView.Visibility = Visibility.Collapsed;
        OverviewView.Visibility = Visibility.Visible;
        BuildCommandList();
    }

    // ---- running ----

    private async void OnRun(object sender, RoutedEventArgs e)
    {
        if (command is not { } c)
        {
            return;
        }
        var args = FormArgs();
        var problems = c.Check(args, [pickedKind]);
        FormProblems.Text = string.Join("\n", problems);
        FormProblems.Visibility = problems.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        if (problems.Count > 0)
        {
            return;
        }
        var kind = pickedKind;
        var targets = Targets(kind);
        if (targets.Count == 0)
        {
            return;
        }
        var what = TargetsText(kind, targets.Count);
        // with every change confirmed, each target asks in the results instead
        var ask = settings.ConfirmChanges && !c.ReadOnly;
        if (!ask && (c.Destructive || targets.Count > c.ConfirmAbove))
        {
            var names = string.Join("\n", targets.Take(8).Select(t => $"• {t.Title}")) + (targets.Count > 8 ? $"\n…and {targets.Count - 8:N0} more" : "");
            var inputs = InputsText(c, args);
            var body = c.Description + (inputs.Length > 0 ? $"\n\n{inputs}" : "") + (kind == NodeKind.Tenant ? "" : $"\n\nOn {what}:\n{names}");
            if (!await ConfirmAsync($"{c.Title}?", body, c.ReadOnly ? "Open them" : "Run", danger: c.Destructive))
            {
                return;
            }
        }
        if (StartJob(c, targets, args, ask) is { } job)
        {
            WatchJob(job);
        }
    }

    /// <summary>The inputs given, one a line ("Show in search: off"), for confirming.</summary>
    private static string InputsText(Command c, CommandArgs args) => string.Join("\n", c.Parameters
        .Where(p => !string.IsNullOrWhiteSpace(args[p.Name]))
        .Select(p => $"{p.Label}: {p.Kind switch
        {
            ParameterKind.Toggle => args[p.Name] == "true" ? "on" : "off",
            ParameterKind.Choice => p.Choices.FirstOrDefault(x => x.Value == args[p.Name])?.Label ?? args[p.Name],
            _ => args[p.Name]!.Trim().ReplaceLineEndings(" "),
        }}"));

    /// <summary>The pane's results follow a job: its rows, its progress, and Pause, Cancel and Run all for it.</summary>
    private void WatchJob(Job job)
    {
        Unwatch();
        watchedJob = job;
        paneBinding = new JobBinding(job, ResultsList, UpdateWatched);
        FormView.Visibility = Visibility.Collapsed;
        OverviewView.Visibility = Visibility.Collapsed;
        PlanView.Visibility = Visibility.Collapsed;
        ResultsView.Visibility = Visibility.Visible;
        UpdateWatched();
    }

    private void Unwatch()
    {
        paneBinding?.Dispose();
        paneBinding = null;
        watchedJob = null;
    }

    private void UpdateWatched()
    {
        if (watchedJob is not { } job)
        {
            return;
        }
        ResultsTitle.Text = job.Plan is null ? job.Command.Title : $"{job.Plan.Name} · {job.StepTitle}";
        var inputs = InputsText(job.Command, job.Args);
        ResultsInputs.Text = inputs;
        ResultsInputs.Visibility = inputs.Length > 0 ? Visibility.Visible : Visibility.Collapsed;
        ResultsSummary.Text = JobSummary(job);
        ResultsBar.Maximum = Math.Max(1, job.Targets.Count);
        ResultsBar.Value = job.DoneCount;
        ResultsBar.ShowPaused = job.State == JobState.Paused;
        ResultsBar.Visibility = job.IsOver ? Visibility.Collapsed : Visibility.Visible;
        PauseRunButton.Visibility = job.IsOver ? Visibility.Collapsed : Visibility.Visible;
        PauseRunButton.Content = job.State == JobState.Paused ? "Carry on" : "Pause";
        PauseRunButton.IsEnabled = !job.Pausing;
        CancelRunButton.Visibility = job.IsOver ? Visibility.Collapsed : Visibility.Visible;
        RunAllButton.Visibility = job.AskingCount > 0 ? Visibility.Visible : Visibility.Collapsed;
        DoneButton.Visibility = Visibility.Visible;
        DoneButton.Content = job.IsOver ? "Done" : "Close";
        ToolTipService.SetToolTip(DoneButton, job.IsOver ? null : "It carries on in the background: Jobs shows it");
        var failed = job.Targets.Values.Count(t => t.Result is { Ok: false });
        SelectFailedButton.Visibility = job.IsOver && failed > 0 && failed < job.Targets.Count && job.Kind != NodeKind.Tenant ? Visibility.Visible : Visibility.Collapsed;
        if (shownPlan is not null && PlanView.Visibility == Visibility.Visible)
        {
            RefreshPlanCheck();
        }
    }

    /// <summary>Runs every target still waiting to be confirmed.</summary>
    private void OnRunAll(object sender, RoutedEventArgs e) => paneBinding?.RunAll();

    private void OnPauseRun(object sender, RoutedEventArgs e)
    {
        if (watchedJob is not { } job)
        {
            return;
        }
        if (job.State == JobState.Paused)
        {
            _ = ResumeJobAsync(job);
        }
        else
        {
            PauseJob(job);
        }
        UpdateWatched();
    }

    /// <summary>"/sites/hr", or the folder a file is in: where a target is, so same-named ones can be told apart.</summary>
    private static string? WhereText(SpNode n)
    {
        static string PathOf(string url) => Uri.TryCreate(url, UriKind.Absolute, out var u)
            ? u.AbsolutePath.TrimEnd('/') is { Length: > 0 } path ? Uri.UnescapeDataString(path) : u.Host
            : url;
        return n.Kind switch
        {
            NodeKind.Tenant => null,
            NodeKind.SiteCollection or NodeKind.Subsite => PathOf(n.Url),
            NodeKind.Folder or NodeKind.File => n.ParentFolder ?? PathOf(n.WebUrl),
            _ => PathOf(n.WebUrl),
        };
    }

    /// <summary>One target's row in the results: waiting (or asking), running, then how it went.</summary>
    private sealed class RunRow
    {
        private readonly Grid icon = new() { Width = 16, Height = 16, VerticalAlignment = VerticalAlignment.Top, Margin = new Thickness(0, 2, 0, 0) };
        private readonly TextBlock message;
        private readonly StackPanel text = new() { Spacing = 2 };
        private readonly StackPanel buttons = new() { Orientation = Orientation.Horizontal, Spacing = 8, VerticalAlignment = VerticalAlignment.Center, Visibility = Visibility.Collapsed };

        /// <summary>The message always takes two lines, and the buttons sit in them: rows never change height as a run goes.</summary>
        private const int MessageLines = 2;
        private readonly bool danger;
        private readonly Action asking;
        private TaskCompletionSource<bool>? answer;

        public RunRow(SpNode target, bool tenant, bool danger, Action asking)
        {
            this.danger = danger;
            this.asking = asking;
            var grid = new Grid { ColumnSpacing = 10, Padding = new Thickness(12, 8, 12, 8) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.Children.Add(icon);
            if (!tenant)
            {
                text.Children.Add(new TextBlock { Text = target.Title, TextTrimming = TextTrimming.CharacterEllipsis });
                if (WhereText(target) is { } where)
                {
                    text.Children.Add(new TextBlock
                    {
                        Text = where,
                        Style = (Style)Application.Current.Resources["CaptionTextBlockStyle"],
                        Foreground = (Brush)Application.Current.Resources["TextFillColorTertiaryBrush"],
                        TextTrimming = TextTrimming.CharacterEllipsis,
                    });
                }
            }
            var lineHeight = tenant ? 20 : 16;
            message = new TextBlock
            {
                Style = (Style)Application.Current.Resources[tenant ? "BodyTextBlockStyle" : "CaptionSecondary"],
                IsTextSelectionEnabled = true,
                TextWrapping = TextWrapping.Wrap,
                TextTrimming = TextTrimming.CharacterEllipsis,
                MaxLines = MessageLines,
                LineHeight = lineHeight,
                LineStackingStrategy = LineStackingStrategy.BlockLineHeight,
            };
            var slot = new Grid { Height = Math.Max(MessageLines * lineHeight, 32) };
            slot.Children.Add(message);
            slot.Children.Add(buttons);
            text.Children.Add(slot);
            var run = new Button { Content = "Run", MinWidth = 72 };
            if (!danger)
            {
                run.Style = (Style)Application.Current.Resources["AccentButtonStyle"];
            }
            run.Click += (_, _) => Answer(true);
            var skip = new Button { Content = "Skip", MinWidth = 72 };
            skip.Click += (_, _) => Answer(false);
            buttons.Children.Add(run);
            buttons.Children.Add(skip);
            Grid.SetColumn(text, 1);
            grid.Children.Add(text);
            View = grid;
            Waiting();
        }

        public FrameworkElement View { get; }

        public bool IsAsking => answer is not null;

        /// <summary>Waits for its turn.</summary>
        public void Waiting() => Set(Glyph("\uE823", "TextFillColorTertiaryBrush"), "Waiting…");

        /// <summary>Asks whether to run it; <paramref name="tcs"/> gets the answer.</summary>
        public void Ask(TaskCompletionSource<bool> tcs)
        {
            if (tcs.Task.IsCompleted)
            {
                return;
            }
            answer = tcs;
            Set(Glyph("\uE897", danger ? "SystemFillColorCautionBrush" : "SystemFillColorAttentionBrush"), "Run it here?");
            ShowButtons(true);
            asking();
        }

        public void Answer(bool yes)
        {
            if (answer is not { } tcs)
            {
                return;
            }
            answer = null;
            ShowButtons(false);
            if (yes)
            {
                Waiting();
            }
            tcs.TrySetResult(yes);
            asking();
        }

        public void Running() => Set(new ProgressRing { IsActive = true, Width = 16, Height = 16 }, "Running…");

        public void Show(CommandResult r)
        {
            if (answer is not null)
            {
                answer = null;
                ShowButtons(false);
                asking();
            }
            var (glyph, brush) = !r.Ok
                ? ("\uEA39", "SystemFillColorCriticalBrush")
                : r.WasDeclined
                    ? ("\uE711", "TextFillColorTertiaryBrush")
                    : r.WasSkipped
                        ? ("\uE946", "SystemFillColorAttentionBrush")
                        : ("\uE930", "SystemFillColorSuccessBrush");
            Set(Glyph(glyph, brush), r.Message);
            if (r.Detail is { Length: > 0 } detail && text.Children.OfType<Border>().FirstOrDefault() is null)
            {
                text.Children.Add(new Border
                {
                    Background = (Brush)Application.Current.Resources["SubtleFillColorSecondaryBrush"],
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(8, 6, 8, 6),
                    Margin = new Thickness(0, 4, 0, 0),
                    Child = new TextBlock { Text = detail, Style = (Style)Application.Current.Resources["Code"] },
                });
            }
        }

        private static FontIcon Glyph(string glyph, string brush) => new()
        {
            Glyph = glyph,
            FontSize = 16,
            Foreground = (Brush)Application.Current.Resources[brush],
        };

        /// <summary>Run and Skip, in the message's place.</summary>
        private void ShowButtons(bool show)
        {
            buttons.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            message.Visibility = show ? Visibility.Collapsed : Visibility.Visible;
        }

        /// <summary>The status and its message (all of it on hover, when it's cut short).</summary>
        private void Set(UIElement status, string text)
        {
            icon.Children.Clear();
            icon.Children.Add(status);
            message.Text = text;
            ToolTipService.SetToolTip(message, text.Length > 60 ? text : null);
        }
    }

    private void OnCancelRun(object sender, RoutedEventArgs e)
    {
        if (watchedJob is { IsOver: false } job)
        {
            CancelJob(job);
            ResultsSummary.Text = "Stopping after the ones already running…";
        }
    }

    private void OnResultsDone(object sender, RoutedEventArgs e)
    {
        var job = watchedJob;
        Unwatch();
        if (job is { IsOver: false })
        {
            ShowToast($"“{job.Title}” carries on", "It runs in the background: Jobs shows how it's going.", ("Jobs", () => Nav.SelectedItem = HistoryItem));
        }
        // after a plan's run: back to the plan
        if (shownPlan is { } plan)
        {
            ShowPlan(plan);
            return;
        }
        ShowOverview();
    }

    /// <summary>What the watched job has done so far, as results.</summary>
    private List<TargetResult> WatchedResults() =>
        watchedJob?.Targets.Values.Where(t => t.Result is not null).Select(t => new TargetResult(t.Node, t.Result!)).ToList() ?? [];

    private void OnCopyResults(object sender, RoutedEventArgs e)
    {
        var text = new StringBuilder();
        text.AppendLine(ResultsTitle.Text);
        foreach (var r in WatchedResults())
        {
            text.AppendLine($"{(r.Result.Ok ? r.Result.WasSkipped ? "–" : "✓" : "✗")} {r.Target.Url}: {r.Result.Message}");
            if (r.Result.Detail is { Length: > 0 } d)
            {
                text.AppendLine(string.Join(Environment.NewLine, d.Split('\n').Select(l => "    " + l.TrimEnd('\r'))));
            }
        }
        CopyText(text.ToString());
    }

    /// <summary>Ticks only the targets that failed, to try them again.</summary>
    private void OnSelectFailed(object sender, RoutedEventArgs e)
    {
        var failed = WatchedResults().Where(r => !r.Result.Ok).Select(r => r.Target).ToList();
        Unwatch();
        foreach (var key in selection.Keys.ToList())
        {
            Untick(key);
        }
        foreach (var node in failed)
        {
            selection[node.Key] = node;
            foreach (var row in rows.Where(r => r.Node?.Key == node.Key))
            {
                row.SetChecked(true);
            }
        }
        ShowOverview();
        UpdatePane();
        if (command is not null)
        {
            OpenForm(command);
        }
    }

    // ---- showing and hiding ----

    private void OnPaneToggle(object sender, RoutedEventArgs e) => ShowPane(PaneToggle.IsChecked == true);

    private void OnHidePane(object sender, RoutedEventArgs e) => ShowPane(false);

    private void ShowPane(bool show)
    {
        PaneToggle.IsChecked = show;
        Pane.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
        PaneColumn.Width = show ? new GridLength(380) : new GridLength(0);
    }
}
