using System.Diagnostics;
using System.Text;

using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.History;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.App;

/// <summary>
/// What keeps work moving between runs: each connection keeps its own ticks,
/// a command offers the inputs it last ran with, results go to Excel as CSV,
/// and the pane can be made wider for them.
/// </summary>
public sealed partial class MainWindow
{
    // ---- each connection keeps its ticks ----

    private sealed record Ticks(List<SpNode> Items, Dictionary<string, SpNode> SiteOf, Dictionary<string, SpNode> ParentOf, NodeKind Kind);

    /// <summary>What was ticked in each connection when another was picked.</summary>
    private readonly Dictionary<string, Ticks> ticksOf = [];

    /// <summary>Before going to connection <paramref name="next"/>: keeps what's ticked in this one.</summary>
    private void StashTicks(string next)
    {
        if (active is null || active.Id == next)
        {
            return;
        }
        if (selection.Count == 0)
        {
            ticksOf.Remove(active.Id);
            return;
        }
        ticksOf[active.Id] = new Ticks(selection.Values.ToList(), new(siteOf), new(parentOf), pickedKind);
    }

    /// <summary>A connection picked again: what was ticked in it is ticked again (rows show it as they load).</summary>
    private void RestoreTicks(string id)
    {
        if (!ticksOf.Remove(id, out var ticks))
        {
            return;
        }
        foreach (var n in ticks.Items)
        {
            selection[n.Key] = n;
        }
        foreach (var (key, site) in ticks.SiteOf)
        {
            siteOf[key] = site;
        }
        foreach (var (key, parent) in ticks.ParentOf)
        {
            parentOf[key] = parent;
        }
        pickedKind = ticks.Kind;
        UpdatePane();
    }

    // ---- the inputs a command last ran with ----

    /// <summary>Each command's newest run with inputs, read from the history once.</summary>
    private Dictionary<string, RunRecord>? lastRuns;

    private RunRecord? lastInputs;

    private RunRecord? LastRunOf(string commandId)
    {
        lastRuns ??= historyStore.List()
            .Where(r => r.Args.Count > 0)
            .GroupBy(r => r.CommandId)
            .ToDictionary(g => g.Key, g => g.MaxBy(r => r.Started)!);
        return lastRuns.GetValueOrDefault(commandId);
    }

    /// <summary>A link to the inputs the command last ran with, when they're not what the form has.</summary>
    private void OfferLastInputs(Command c, CommandArgs? given)
    {
        lastInputs = null;
        LastInputsLink.Visibility = Visibility.Collapsed;
        if (given is not null || LastRunOf(c.Id) is not { } last)
        {
            return;
        }
        var now = FormArgs();
        if (last.Args.All(a => now[a.Key] == a.Value))
        {
            return;
        }
        lastInputs = last;
        LastInputsLink.Content = new TextBlock { Text = $"Use the inputs of its last run ({When(last.Started)})", TextWrapping = TextWrapping.Wrap };
        ToolTipService.SetToolTip(LastInputsLink, last.Inputs.Count > 0 ? string.Join("\n", last.Inputs) : null);
        LastInputsLink.Visibility = Visibility.Visible;
    }

    private void OnUseLastInputs(object sender, RoutedEventArgs e)
    {
        if (command is not { } c || lastInputs is not { } last)
        {
            return;
        }
        var args = new CommandArgs();
        foreach (var (name, value) in last.Args)
        {
            args[name] = value;
        }
        OpenForm(c, args);
    }

    // ---- results for Excel ----

    private void OnSaveResultsCsv(object sender, RoutedEventArgs e)
    {
        if (watchedJob is { } job)
        {
            _ = SaveCsvAsync(job.Title, LiveLines(job));
        }
    }

    private void OnSaveRunCsv(object sender, RoutedEventArgs e)
    {
        if (shownEntry is not { } entry)
        {
            return;
        }
        var lines = entry.Live is { } job
            ? LiveLines(job)
            : entry.Records.OrderBy(r => r.Step ?? 1)
                .SelectMany(r => r.Targets.Select(t => new ResultLine(r.Step, t.Target, t.Outcome, t.Message, t.Detail, t.Table)))
                .ToList();
        _ = SaveCsvAsync(entry.Title, lines);
    }

    private static List<ResultLine> LiveLines(Job job) => job.Targets.Values
        .Select(t => new ResultLine(job.Plan is null ? null : job.Step, t.Node,
            t.Result is { } r ? RunTarget.OutcomeOf(r) : RunOutcome.Pending, t.Result?.Message ?? "", t.Result?.Detail, t.Result?.Table))
        .ToList();

    /// <summary>Saves results as a CSV (with a BOM, so Excel reads it as UTF-8), then offers to open it.</summary>
    private async Task SaveCsvAsync(string title, IReadOnlyList<ResultLine> lines)
    {
        var name = string.Concat($"{title} {DateTime.Now:yyyy-MM-dd}".Select(ch => Path.GetInvalidFileNameChars().Contains(ch) ? '-' : ch));
        if (await PickSaveFileAsync(name, "CSV for Excel", ".csv") is not { } path)
        {
            return;
        }
        try
        {
            await File.WriteAllTextAsync(path, ResultsCsv.Build(lines), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            ShowError(new InvalidOperationException($"Couldn't save the results: {ex.Message}"));
            return;
        }
        ShowToast($"Saved {Path.GetFileName(path)}", $"{Plural(lines.Count, "result", "results")}, a row for each line they list.", ("Open", () =>
        {
            try
            {
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true })?.Dispose();
            }
            catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException)
            {
                ShowError(new InvalidOperationException($"Couldn't open it: {ex.Message}"));
            }
        }));
    }

    // ---- the pane's width ----

    private double dragFrom;

    private void SetUpSplitter()
    {
        PaneSplitter.Started += () => dragFrom = PaneColumn.ActualWidth;
        PaneSplitter.Dragged += dx => PaneColumn.Width = new GridLength(ClampPane(dragFrom - dx));
        PaneSplitter.Released += () => SavePaneWidth(PaneColumn.ActualWidth);
        PaneSplitter.Reset += () =>
        {
            PaneColumn.Width = new GridLength(Settings.AppSettings.DefaultPaneWidth);
            SavePaneWidth(Settings.AppSettings.DefaultPaneWidth);
        };
    }

    /// <summary>At least 320 wide, and the table keeps at least 440.</summary>
    private double ClampPane(double width) =>
        Math.Clamp(width, 320, Math.Max(320, WorkArea.ActualWidth - PaneSplitter.ActualWidth - 440));

    private void SavePaneWidth(double width)
    {
        if (Math.Abs(width - settings.PaneWidth) < 1)
        {
            return;
        }
        settings = settings with { PaneWidth = Math.Round(width) };
        try
        {
            settingsStore.Save(settings);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Diagnostics.Log.Write("app", $"couldn't save the pane's width: {ex.Message}");
        }
    }
}
