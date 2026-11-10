using System.Text;

using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.History;
using ScriptlessSharePoint.Selections;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.App;

/// <summary>
/// The history: every command run from the app, newest first, kept in
/// History\ in the data folder (Core\History). A run is saved as it starts
/// and again as it ends. From one, its targets can be ticked again (all, or
/// the failed ones) and it can be run again with the same inputs.
/// </summary>
public sealed partial class MainWindow
{
    private const string HistoryTag = "history";

    private readonly RunHistoryStore historyStore = new();

    /// <summary>The runs the page lists, as last read.</summary>
    private List<RunRecord> history = [];

    /// <summary>The run the page shows.</summary>
    private RunRecord? shownRun;

    /// <summary>A command to open, with the inputs it had, once the table shows a run's targets.</summary>
    private (Command Command, CommandArgs Args)? pendingForm;

    /// <summary>Saves a run, keeping it as it was when the disk says no: history never stops a run.</summary>
    private RunRecord SaveRun(RunRecord run)
    {
        try
        {
            return historyStore.Save(run);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            Diagnostics.Log.Write("history", $"couldn't save {run.CommandId}: {e.Message}");
            return run;
        }
    }

    // ---- the page ----

    private void ShowHistory()
    {
        ShowPage(HistoryPage);
        LoadHistory(shownRun?.File);
    }

    private void LoadHistory(string? select)
    {
        history = historyStore.List().ToList();
        HistoryList.Items.Clear();
        var connections = workspace.Load().Connections.ToDictionary(c => c.Id, c => c.Name);
        foreach (var r in history)
        {
            HistoryList.Items.Add(HistoryEntry(r, connections.GetValueOrDefault(r.ConnectionId) ?? r.ConnectionName));
        }
        HistoryEmpty.Visibility = history.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        ClearHistoryButton.IsEnabled = history.Count > 0;
        var index = history.FindIndex(r => r.File == select);
        HistoryList.SelectedIndex = index >= 0 ? index : history.Count > 0 ? 0 : -1;
        if (history.Count == 0)
        {
            ShowRun(null);
        }
    }

    private static FrameworkElement HistoryEntry(RunRecord r, string connection)
    {
        var (glyph, brush) = Outcome(r);
        var grid = new Grid { ColumnSpacing = 12, Padding = new Thickness(4, 8, 4, 8) };
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
        text.Children.Add(new TextBlock { Text = r.Title, TextTrimming = TextTrimming.CharacterEllipsis });
        text.Children.Add(new TextBlock
        {
            Text = $"{TargetsText(r.Kind, r.Targets.Count)} · {RunSummary(r)}",
            Style = (Style)Application.Current.Resources["CaptionSecondary"],
            TextTrimming = TextTrimming.CharacterEllipsis,
        });
        text.Children.Add(new TextBlock
        {
            Text = $"{When(r.Started)} · {connection}",
            Style = (Style)Application.Current.Resources["CaptionTextBlockStyle"],
            Foreground = (Brush)Application.Current.Resources["TextFillColorTertiaryBrush"],
            TextTrimming = TextTrimming.CharacterEllipsis,
        });
        Grid.SetColumn(text, 1);
        grid.Children.Add(text);
        return grid;
    }

    /// <summary>The run's icon: failed (any target), didn't finish, all well.</summary>
    private static (string Glyph, string Brush) Outcome(RunRecord r) =>
        r.Count(RunOutcome.Failed) > 0 ? ("", "SystemFillColorCriticalBrush")
        : r.Finished is null || r.Cancelled ? ("", "SystemFillColorAttentionBrush")
        : ("", "SystemFillColorSuccessBrush");

    /// <summary>"2 done · 1 failed", "Cancelled · 3 done", "Didn't finish".</summary>
    private static string RunSummary(RunRecord r) => string.Join(" · ", new[]
    {
        r.Finished is null ? "Didn't finish" : r.Cancelled ? "Cancelled" : null,
        r.Count(RunOutcome.Done) is > 0 and var done ? $"{done:N0} done" : null,
        r.Count(RunOutcome.Unchanged) is > 0 and var same ? $"{same:N0} unchanged" : null,
        r.Count(RunOutcome.NotRun) is > 0 and var skipped ? $"{skipped:N0} not run" : null,
        r.Count(RunOutcome.Failed) is > 0 and var failed ? $"{failed:N0} failed" : null,
    }.OfType<string>().DefaultIfEmpty("Nothing run"));

    /// <summary>"Today 12:05", "Yesterday 09:14", "3 Nov 2026 16:40".</summary>
    private static string When(DateTimeOffset at)
    {
        var local = at.LocalDateTime;
        var day = (DateTime.Today - local.Date).Days switch
        {
            0 => "Today",
            1 => "Yesterday",
            _ => local.ToString("d MMM yyyy", System.Globalization.CultureInfo.CurrentCulture),
        };
        return $"{day} {local:HH:mm}";
    }

    private static string Took(TimeSpan t) =>
        t.TotalSeconds < 1 ? "under a second" : t.TotalMinutes < 1 ? $"{t.TotalSeconds:0} s" : t.TotalHours < 1 ? $"{(int)t.TotalMinutes} min {t.Seconds} s" : $"{(int)t.TotalHours} h {t.Minutes} min";

    private void OnHistoryPicked(object sender, SelectionChangedEventArgs e) =>
        ShowRun(HistoryList.SelectedIndex is >= 0 and var i && i < history.Count ? history[i] : null);

    private void ShowRun(RunRecord? r)
    {
        shownRun = r;
        RunDetail.Visibility = r is null ? Visibility.Collapsed : Visibility.Visible;
        RunTargets.Items.Clear();
        if (r is null)
        {
            return;
        }
        var connection = workspace.Load().Find(r.ConnectionId);
        RunTitle.Text = r.Title;
        RunWhen.Text = $"{When(r.Started)}{(r.Finished is { } end ? $" · took {Took(end - r.Started)}" : "")} · {connection?.Name ?? $"{r.ConnectionName} (connection removed)"}";
        RunInputs.Text = string.Join("\n", r.Inputs);
        RunInputs.Visibility = r.Inputs.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        RunSummaryText.Text = $"On {TargetsText(r.Kind, r.Targets.Count)} · {RunSummary(r)}";

        var tenant = r.Kind == NodeKind.Tenant;
        var hasCommand = CommandCatalog.Find(r.CommandId) is not null;
        RunAgainButton.IsEnabled = connection is not null && hasCommand && running is null;
        ToolTipService.SetToolTip(RunAgainButton, !hasCommand
            ? "This command isn't in the app any more."
            : tenant ? "Opens the command with the same inputs." : "Ticks the same targets and opens the command with the same inputs.");
        TickRunButton.Visibility = tenant ? Visibility.Collapsed : Visibility.Visible;
        TickRunButton.IsEnabled = connection is not null;
        var failed = r.Targets.Count(t => t.Outcome is RunOutcome.Failed or RunOutcome.Pending);
        TickRunFailedButton.Visibility = !tenant && failed > 0 && failed < r.Targets.Count ? Visibility.Visible : Visibility.Collapsed;
        TickRunFailedButton.IsEnabled = connection is not null;

        foreach (var t in r.Targets)
        {
            var row = new RunRow(t.Target, tenant, false, () => { });
            row.Show(t.Outcome switch
            {
                RunOutcome.Done => CommandResult.Done(t.Message, t.Detail),
                RunOutcome.Unchanged => CommandResult.Skipped(t.Message),
                RunOutcome.NotRun => CommandResult.Declined() with { Message = t.Message },
                _ => CommandResult.Failed(t.Message) with { Detail = t.Detail },
            });
            RunTargets.Items.Add(row.View);
        }
    }

    // ---- from a run ----

    private void OnRunAgain(object sender, RoutedEventArgs e)
    {
        if (shownRun is { } r && CommandCatalog.Find(r.CommandId) is { } c)
        {
            _ = OpenRunAsync(r, failedOnly: false, c);
        }
    }

    private void OnTickRun(object sender, RoutedEventArgs e)
    {
        if (shownRun is { } r)
        {
            _ = OpenRunAsync(r, failedOnly: false, null);
        }
    }

    private void OnTickRunFailed(object sender, RoutedEventArgs e)
    {
        if (shownRun is { } r)
        {
            _ = OpenRunAsync(r, failedOnly: true, null);
        }
    }

    /// <summary>
    /// Shows a run's targets ticked in its connection's table (as a saved
    /// selection shows: each under the row it was ticked in, nothing loaded),
    /// then, with <paramref name="again"/>, opens that command with the same inputs.
    /// </summary>
    private async Task OpenRunAsync(RunRecord r, bool failedOnly, Command? again)
    {
        if (workspace.Load().Find(r.ConnectionId) is null || running is not null)
        {
            return;
        }
        var args = new CommandArgs();
        foreach (var (name, value) in r.Args)
        {
            args[name] = value;
        }
        if (r.Kind == NodeKind.Tenant)
        {
            // nothing to tick: the connection, and the command for the tenant
            SelectNavQuietly(r.ConnectionId);
            await ShowConnectionAsync(r.ConnectionId);
            if (again is not null && running is null)
            {
                ShowPane(true);
                pickedKind = NodeKind.Tenant;
                UpdatePane();
                OpenForm(again, args);
            }
            return;
        }

        var targets = r.Targets.Where(t => !failedOnly || t.Outcome is RunOutcome.Failed or RunOutcome.Pending).ToList();
        var s = new SavedSelection
        {
            Name = $"{(failedOnly ? "Failed in " : "")}{r.Title}, {When(r.Started)}",
            ConnectionId = r.ConnectionId,
            Saved = r.Started,
            Items = targets.Select(t => t.Target).ToList(),
            Sites = targets.Select(t => t.Site).OfType<SpNode>().DistinctBy(n => n.Key).ToList(),
            Parents = targets.Select(t => t.Parent).OfType<SpNode>().DistinctBy(n => n.Key).ToList(),
        };
        pendingForm = again is null ? null : (again, args);
        if (active?.Id != r.ConnectionId || identity is null)
        {
            pendingSaved = s;
            SelectNavQuietly(r.ConnectionId);
            await ShowConnectionAsync(r.ConnectionId, force: active?.Id == r.ConnectionId);
            return;
        }
        SelectNavQuietly(r.ConnectionId);
        ShowSaved(s);
    }

    /// <summary>After a run's targets show: the command it ran, with its inputs, when asked for.</summary>
    private void OpenPendingForm()
    {
        if (pendingForm is { } f && running is null)
        {
            pendingForm = null;
            if (f.Command.AppliesTo(pickedKind) && Targets(pickedKind).Count > 0)
            {
                OpenForm(f.Command, f.Args);
            }
        }
        pendingForm = null;
    }

    private void OnCopyRun(object sender, RoutedEventArgs e)
    {
        if (shownRun is not { } r)
        {
            return;
        }
        var text = new StringBuilder();
        text.AppendLine($"{r.Title} · {r.Started.LocalDateTime:g} · {r.ConnectionName}");
        foreach (var line in r.Inputs)
        {
            text.AppendLine($"  {line}");
        }
        foreach (var t in r.Targets)
        {
            var mark = t.Outcome switch { RunOutcome.Done => "✓", RunOutcome.Failed or RunOutcome.Pending => "✗", _ => "–" };
            text.AppendLine($"{mark} {t.Target.Url}: {t.Message}");
            if (t.Detail is { Length: > 0 } d)
            {
                text.AppendLine(string.Join(Environment.NewLine, d.Split('\n').Select(l => "    " + l.TrimEnd('\r'))));
            }
        }
        CopyText(text.ToString());
    }

    private void OnDeleteRun(object sender, RoutedEventArgs e)
    {
        if (shownRun is not { } r)
        {
            return;
        }
        var i = history.FindIndex(x => x.File == r.File);
        try
        {
            historyStore.Delete(r.File);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return;
        }
        shownRun = null;
        LoadHistory(i >= 0 && i + 1 < history.Count ? history[i + 1].File : null);
    }

    private async void OnClearHistory(object sender, RoutedEventArgs e)
    {
        if (!await ConfirmAsync("Clear the history?", "Every run goes from the list. Nothing changes in SharePoint.", "Clear", danger: false))
        {
            return;
        }
        try
        {
            historyStore.Clear();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
        }
        shownRun = null;
        LoadHistory(null);
    }
}
