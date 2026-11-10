using System.Globalization;
using System.Text;

using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.Connections;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.App;

/// <summary>
/// The pane beside the table, as File Size's Selected pane: what's ticked, and
/// the commands for one kind of it (or for the tenant). A command opens its
/// inputs here; Run runs it on every ticked thing of that kind, and the
/// results stay until Done.
/// </summary>
public sealed partial class MainWindow
{
    /// <summary>The kind the pane acts on; Tenant when nothing's ticked.</summary>
    private NodeKind pickedKind = NodeKind.Tenant;
    private bool fillingKinds;
    private Command? command;
    private readonly List<(CommandParameter Parameter, Func<string?> Value)> fields = [];
    private CancellationTokenSource? running;
    private List<TargetResult> results = [];

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
    }

    // ---- a command's inputs ----

    private void OpenForm(Command c)
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
        RunButton.Content = c.ReadOnly ? "Show" : "Run";

        FormFields.Children.Clear();
        fields.Clear();
        var defaults = c.Defaults();
        foreach (var p in c.Parameters)
        {
            var (control, value) = Field(p, defaults[p.Name]);
            FormFields.Children.Add(control);
            fields.Add((p, value));
        }

        OverviewView.Visibility = Visibility.Collapsed;
        ResultsView.Visibility = Visibility.Collapsed;
        FormView.Visibility = Visibility.Visible;
        if (FormFields.Children.FirstOrDefault() is Control first)
        {
            first.Focus(FocusState.Programmatic);
        }
    }

    /// <summary>The control for one input, and how to read it.</summary>
    private (FrameworkElement Control, Func<string?> Value) Field(CommandParameter p, string? value)
    {
        switch (p.Kind)
        {
            case ParameterKind.Toggle:
            {
                var toggle = new ToggleSwitch { Header = p.Label, IsOn = value == "true" };
                return (WithHint(toggle, p.Hint), () => toggle.IsOn ? "true" : "false");
            }
            case ParameterKind.Choice:
            {
                var box = new ComboBox { Header = p.Label, HorizontalAlignment = HorizontalAlignment.Stretch, Description = p.Hint };
                foreach (var choice in p.Choices)
                {
                    box.Items.Add(choice.Label);
                }
                box.SelectedIndex = Math.Max(0, p.Choices.ToList().FindIndex(x => x.Value == value));
                return (box, () => box.SelectedIndex >= 0 ? p.Choices[box.SelectedIndex].Value : null);
            }
            case ParameterKind.Number:
            {
                var number = new NumberBox
                {
                    Header = p.Label,
                    Description = p.Hint,
                    PlaceholderText = p.Placeholder ?? "",
                    Value = long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : double.NaN,
                    SpinButtonPlacementMode = NumberBoxSpinButtonPlacementMode.Hidden,
                    ValidationMode = NumberBoxValidationMode.Disabled,
                };
                return (number, () => double.IsNaN(number.Value) ? null : ((long)Math.Round(number.Value)).ToString(CultureInfo.InvariantCulture));
            }
            case ParameterKind.LocalFile or ParameterKind.LocalFolder:
            {
                var box = new TextBox { Header = p.Label, Description = p.Hint, Text = value ?? "", PlaceholderText = p.Kind == ParameterKind.LocalFile ? "Choose a file" : "Choose a folder" };
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
                    Description = p.Hint,
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
        panel.Children.Add(new TextBlock { Text = hint, Style = (Style)Application.Current.Resources["CaptionSecondary"] });
        return panel;
    }

    private void OnBackToCommands(object sender, RoutedEventArgs e) => ShowOverview();

    private void ShowOverview()
    {
        if (running is not null)
        {
            return; // a run's results stay until it's over
        }
        FormView.Visibility = Visibility.Collapsed;
        ResultsView.Visibility = Visibility.Collapsed;
        OverviewView.Visibility = Visibility.Visible;
        BuildCommandList();
    }

    // ---- running ----

    private async void OnRun(object sender, RoutedEventArgs e)
    {
        if (command is not { } c || running is not null)
        {
            return;
        }
        var args = new CommandArgs();
        foreach (var (p, value) in fields)
        {
            args[p.Name] = value();
        }
        var problems = c.Check(args);
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
        if (c.Destructive || targets.Count > c.ConfirmAbove)
        {
            var names = string.Join("\n", targets.Take(8).Select(t => $"• {t.Title}")) + (targets.Count > 8 ? $"\n…and {targets.Count - 8:N0} more" : "");
            var body = kind == NodeKind.Tenant
                ? c.Description
                : $"{c.Description}\n\nOn {what}:\n{names}";
            if (!await ConfirmAsync($"{c.Title}?", body, c.ReadOnly ? "Open them" : "Run", danger: c.Destructive))
            {
                return;
            }
        }
        await RunAsync(c, targets, args, kind);
    }

    private async Task RunAsync(Command c, List<SpNode> targets, CommandArgs args, NodeKind kind)
    {
        Session session;
        try
        {
            session = workspace.Current();
        }
        catch (Exception e)
        {
            ShowError(e);
            return;
        }
        results = [];
        ResultsList.Items.Clear();
        ResultsTitle.Text = c.Title;
        ResultsSummary.Text = $"On {TargetsText(kind, targets.Count)}…";
        ResultsBar.Maximum = targets.Count;
        ResultsBar.Value = 0;
        ResultsBar.Visibility = Visibility.Visible;
        CancelRunButton.Visibility = Visibility.Visible;
        DoneButton.Visibility = Visibility.Collapsed;
        SelectFailedButton.Visibility = Visibility.Collapsed;
        FormView.Visibility = Visibility.Collapsed;
        OverviewView.Visibility = Visibility.Collapsed;
        ResultsView.Visibility = Visibility.Visible;

        running = new CancellationTokenSource();
        // progress arrives posted to this thread, possibly after the run has
        // ended: then the run's own list, shown below, already has it
        var finished = false;
        var progress = new Progress<TargetResult>(r =>
        {
            if (finished)
            {
                return;
            }
            results.Add(r);
            ResultsBar.Value = results.Count;
            ResultsSummary.Text = $"{results.Count:N0} of {targets.Count:N0}";
            ResultsList.Items.Add(ResultRow(r, kind == NodeKind.Tenant));
        });
        IReadOnlyList<TargetResult> all;
        try
        {
            all = await CommandRunner.RunAsync(c, targets, args, session.Client, this, progress, running.Token);
        }
        catch (Exception e)
        {
            all = results;
            ShowError(e);
        }
        finally
        {
            running.Dispose();
            running = null;
        }

        // the run's list is the whole story (with the ones cancelled before they started)
        finished = true;
        results = [.. all];
        ResultsList.Items.Clear();
        foreach (var r in results)
        {
            ResultsList.Items.Add(ResultRow(r, kind == NodeKind.Tenant));
        }
        var done = results.Count(r => r.Result is { Ok: true, WasSkipped: false });
        var skipped = results.Count(r => r.Result.WasSkipped);
        var failed = results.Count(r => !r.Result.Ok);
        ResultsSummary.Text = string.Join(" · ", new[]
        {
            done > 0 ? $"{done:N0} done" : null,
            skipped > 0 ? $"{skipped:N0} unchanged" : null,
            failed > 0 ? $"{failed:N0} failed" : null,
        }.OfType<string>());
        ResultsBar.Visibility = Visibility.Collapsed;
        CancelRunButton.Visibility = Visibility.Collapsed;
        DoneButton.Visibility = Visibility.Visible;
        SelectFailedButton.Visibility = failed > 0 && failed < results.Count && kind != NodeKind.Tenant ? Visibility.Visible : Visibility.Collapsed;
        Apply(c, results);
    }

    private static FrameworkElement ResultRow(TargetResult r, bool tenant)
    {
        var (glyph, brush) = !r.Result.Ok
            ? ("", "SystemFillColorCriticalBrush")
            : r.Result.WasSkipped
                ? ("", "SystemFillColorAttentionBrush")
                : ("", "SystemFillColorSuccessBrush");
        var grid = new Grid { ColumnSpacing = 10, Padding = new Thickness(12, 8, 12, 8) };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.Children.Add(new FontIcon
        {
            Glyph = glyph,
            FontSize = 16,
            Foreground = (Brush)Application.Current.Resources[brush],
            VerticalAlignment = VerticalAlignment.Top,
            Margin = new Thickness(0, 2, 0, 0),
        });
        var text = new StackPanel { Spacing = 2 };
        if (!tenant)
        {
            text.Children.Add(new TextBlock { Text = r.Target.Title, TextTrimming = TextTrimming.CharacterEllipsis });
        }
        text.Children.Add(new TextBlock
        {
            Text = r.Result.Message,
            Style = (Style)Application.Current.Resources[tenant ? "BodyTextBlockStyle" : "CaptionSecondary"],
            IsTextSelectionEnabled = true,
            TextWrapping = TextWrapping.Wrap,
        });
        if (r.Result.Detail is { Length: > 0 } detail)
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
        Grid.SetColumn(text, 1);
        grid.Children.Add(text);
        return grid;
    }

    private void OnCancelRun(object sender, RoutedEventArgs e)
    {
        running?.Cancel();
        CancelRunButton.Visibility = Visibility.Collapsed;
        ResultsSummary.Text = "Stopping after the ones already running…";
    }

    private void OnResultsDone(object sender, RoutedEventArgs e) => ShowOverview();

    private void OnCopyResults(object sender, RoutedEventArgs e)
    {
        var text = new StringBuilder();
        text.AppendLine(ResultsTitle.Text);
        foreach (var r in results)
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
        var failed = results.Where(r => !r.Result.Ok).Select(r => r.Target).ToList();
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
