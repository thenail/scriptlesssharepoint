using System.Globalization;
using System.Text;

using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.History;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.App;

/// <summary>
/// Reports: what a command listed (users, columns, versions, details…) as a
/// table in the site table's place, a row per thing it listed (and, over many
/// targets, which one it's in), sortable by any column, filtered by what's
/// typed, copied for Excel or saved as CSV. Each target's table lines up with
/// the others by column name. A read-only command that lists much opens its
/// report when it's done; any results (and any finished job) can show theirs.
/// </summary>
public sealed partial class MainWindow
{
    /// <summary>What one target gave a report: whether it went well, and its table or lines.</summary>
    private sealed record ReportSource(SpNode Target, bool Ok, ReportTable? Table, string? Detail);

    /// <summary>A row of the report, and the target it came from.</summary>
    private sealed class ReportRow(SpNode target, IReadOnlyList<string> cells)
    {
        public SpNode Target { get; } = target;
        public IReadOnlyList<string> Cells { get; } = cells;
    }

    private static readonly StringComparer ReportOrder = StringComparer.Create(CultureInfo.CurrentCulture, CompareOptions.IgnoreCase | CompareOptions.NumericOrdering);

    private IReadOnlyList<string> reportColumns = [];
    private GridLength[] reportWidths = [];
    private List<ReportRow> reportRows = [];
    private List<ReportRow> reportShown = [];
    private string reportTitle = "";
    private int reportSort = -1;
    private bool reportDescending;

    /// <summary>Bumped when the columns change, so recycled rows lay themselves out again.</summary>
    private int reportVersion;

    /// <summary>The title and status before the report was shown, to put back when it closes.</summary>
    private string[]? crumbsBeforeReport;
    private (string Title, string Text)? statusBeforeReport;

    /// <summary>Whether the pane showed before the report (it makes room for the report, and comes back after).</summary>
    private bool paneBeforeReport;

    private bool reportKeys;

    /// <summary>" from 3 site collections", after the count of rows.</summary>
    private string reportFrom = "";

    /// <summary>"12 rows from 3 site collections", or, filtered, "2 of 12 rows…".</summary>
    private string ReportCount() =>
        (reportShown.Count < reportRows.Count ? $"{reportShown.Count:N0} of " : "") + Plural(reportRows.Count, "row", "rows") + reportFrom;

    // ---- where reports come from ----

    private static List<ReportSource> ReportSourcesOf(Job job) => job.Targets.Values
        .Where(t => t.Result is not null)
        .Select(t => new ReportSource(t.Node, t.Result!.Ok && !t.Result.WasDeclined, t.Result.Table, t.Result.Detail))
        .ToList();

    private static List<ReportSource> ReportSourcesOf(RunRecord r) => r.Targets
        .Select(t => new ReportSource(t.Target, t.Outcome is RunOutcome.Done or RunOutcome.Unchanged, t.Table, t.Detail))
        .ToList();

    private static bool HasReport(IEnumerable<ReportSource> sources) =>
        sources.Any(s => s.Ok && (s.Table is { Rows.Count: > 0 } || !string.IsNullOrWhiteSpace(s.Detail)));

    /// <summary>A finished job's report: its last step that listed anything.</summary>
    private static RunRecord? ReportRecord(IReadOnlyList<RunRecord> records) =>
        records.OrderBy(r => r.Step ?? 1).LastOrDefault(r => HasReport(ReportSourcesOf(r)));

    /// <summary>"List users", or "Plan · Step 2: List columns".</summary>
    private static string ReportName(RunRecord r) => r.PlanName is null ? r.Title : $"{r.PlanName} · step {r.Step}: {r.Title}";

    private void OnShowReport(object sender, RoutedEventArgs e)
    {
        if (watchedJob is { } job)
        {
            ShowReport(job.Plan is null ? job.Command.Title : $"{job.Plan.Name} · {job.StepTitle}", ReportSourcesOf(job));
        }
    }

    /// <summary>From the Jobs page: its connection, then its report in the table's place.</summary>
    private async void OnReportRun(object sender, RoutedEventArgs e)
    {
        if (shownEntry is not { } entry)
        {
            return;
        }
        string connectionId;
        string title;
        List<ReportSource> sources;
        if (entry.Live is { } job)
        {
            (connectionId, title, sources) = (job.ConnectionId, job.Title, ReportSourcesOf(job));
        }
        else if (ReportRecord(entry.Records) is { } record)
        {
            (connectionId, title, sources) = (record.ConnectionId, ReportName(record), ReportSourcesOf(record));
        }
        else
        {
            return;
        }
        if (workspace.Load().Find(connectionId) is null)
        {
            return;
        }
        if (active?.Id != connectionId)
        {
            SelectNavQuietly(connectionId);
            await ShowConnectionAsync(connectionId);
        }
        else
        {
            // the same connection: what its table shows stays picked
            SelectNavForShown();
        }
        ShowReport(title, sources);
    }

    // ---- showing one ----

    private void ShowReport(string title, IReadOnlyList<ReportSource> sources)
    {
        // each target's table (or its lines, split), lined up by column name
        var tables = sources
            .Where(s => s.Ok)
            .Select(s => (s.Target, Table: s.Table ?? (string.IsNullOrWhiteSpace(s.Detail) ? null : ReportTable.FromLines(s.Detail))))
            .Where(x => x.Table is { Rows.Count: > 0 })
            .ToList();
        var columns = ReportTable.Union(tables.Select(x => x.Table!));
        var many = sources.Select(s => s.Target.Key).Distinct().Count() > 1;
        var kind = sources.Count > 0 ? sources[0].Target.Kind : NodeKind.Tenant;
        reportColumns = many ? [Capitalized(NodeKinds.Name(kind)), .. columns] : columns;
        reportRows = tables
            .SelectMany(x => x.Table!.Rows.Select(r => new ReportRow(x.Target, many ? [x.Target.Title, .. x.Table.Cells(r, columns)] : x.Table.Cells(r, columns))))
            .ToList();
        reportTitle = title;
        reportSort = -1;
        reportDescending = false;
        reportVersion++;
        reportWidths = reportColumns.Select((name, i) =>
        {
            var cells = reportRows.Take(300).Select(r => r.Cells[i].Length).DefaultIfEmpty(0);
            return new GridLength(Math.Clamp(Math.Max(name.Length, cells.Average() * 1.2), 6, 40), GridUnitType.Star);
        }).ToArray();

        reportFrom = many ? $" from {TargetsText(kind, tables.Count)}" : "";
        var failed = sources.Count(s => !s.Ok);
        var nothing = sources.Count(s => s.Ok) - tables.Count;

        ReportFilter.Text = "";
        ApplyReport();

        if (ReportCard.Visibility != Visibility.Visible)
        {
            crumbsBeforeReport = Crumbs.ItemsSource as string[] ?? (active is null ? [] : [active.Name]);
            statusBeforeReport = (StatusTitle.Text, StatusText.Text);
            paneBeforeReport = Pane.Visibility == Visibility.Visible;
        }
        // how much it listed, where the table's count was
        SetStatus(ReportCount(), string.Join(" · ", new[]
        {
            nothing > 0 ? $"{nothing:N0} listed nothing" : null,
            failed > 0 ? $"{failed:N0} failed or didn't run (the results say why)" : null,
            "Click a column to sort by it; right-click a row to copy or open it.",
        }.OfType<string>()));
        string[] crumbs = [.. crumbsBeforeReport ?? Array.Empty<string>(), $"Report: {title}"];
        Crumbs.ItemsSource = crumbs;
        TableCard.Visibility = Visibility.Collapsed;
        ReportCard.Visibility = Visibility.Visible;
        ShowPane(false);
        ShowPage(TenantPage);
        if (!reportKeys)
        {
            // Esc, inside the report: back to the table
            var esc = new KeyboardAccelerator { Key = Windows.System.VirtualKey.Escape };
            esc.Invoked += (_, e) =>
            {
                e.Handled = true;
                HideReport();
            };
            ReportCard.KeyboardAccelerators.Add(esc);
            reportKeys = true;
        }
    }

    private static string Capitalized(string s) => s.Length == 0 ? s : char.ToUpper(s[0], CultureInfo.CurrentCulture) + s[1..];

    /// <summary>The table again in the report's place; the title as it was (unless what replaces it sets its own).</summary>
    private void HideReport(bool restoreTitle = true)
    {
        if (ReportCard.Visibility != Visibility.Visible)
        {
            return;
        }
        ReportCard.Visibility = Visibility.Collapsed;
        TableCard.Visibility = Visibility.Visible;
        if (restoreTitle && crumbsBeforeReport is not null)
        {
            Crumbs.ItemsSource = crumbsBeforeReport;
        }
        if (restoreTitle && statusBeforeReport is { } status)
        {
            SetStatus(status.Title, status.Text);
        }
        crumbsBeforeReport = null;
        statusBeforeReport = null;
        if (paneBeforeReport)
        {
            ShowPane(true);
        }
        ReportList.ItemsSource = null;
    }

    private void OnCloseReport(object sender, RoutedEventArgs e) => HideReport();

    // ---- sorting and filtering ----

    private void OnReportFilterChanged(AutoSuggestBox sender, AutoSuggestBoxTextChangedEventArgs args) => ApplyReport();

    /// <summary>The rows that match the filter, in the order picked, and the header to match.</summary>
    private void ApplyReport()
    {
        var filter = ReportFilter.Text?.Trim() ?? "";
        IEnumerable<ReportRow> rows = reportRows.Where(r => filter.Length == 0 || r.Cells.Any(c => c.Contains(filter, StringComparison.CurrentCultureIgnoreCase)));
        if (reportSort >= 0)
        {
            var i = reportSort;
            // empty cells last, whichever way it's sorted
            rows = rows.OrderBy(r => r.Cells[i].Length == 0);
            rows = reportDescending
                ? ((IOrderedEnumerable<ReportRow>)rows).ThenByDescending(r => r.Cells[i], ReportOrder)
                : ((IOrderedEnumerable<ReportRow>)rows).ThenBy(r => r.Cells[i], ReportOrder);
        }
        reportShown = rows.ToList();
        ReportList.ItemsSource = reportShown;
        ReportEmpty.Text = reportRows.Count == 0 ? "It listed nothing." : "No rows match the filter.";
        ReportEmpty.Visibility = reportShown.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        if (ReportCard.Visibility == Visibility.Visible)
        {
            StatusTitle.Text = ReportCount();
        }
        BuildReportHeader();
    }

    /// <summary>A header per column: click to sort by it, again to reverse.</summary>
    private void BuildReportHeader()
    {
        ReportHeader.Children.Clear();
        ReportHeader.ColumnDefinitions.Clear();
        for (var i = 0; i < reportColumns.Count; i++)
        {
            ReportHeader.ColumnDefinitions.Add(new ColumnDefinition { Width = reportWidths[i], MinWidth = 44 });
            var text = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
            text.Children.Add(new TextBlock { Text = reportColumns[i], Style = (Style)Application.Current.Resources["ColumnHeader"], Margin = new Thickness(0) });
            if (reportSort == i)
            {
                text.Children.Add(new FontIcon { Glyph = reportDescending ? "" : "", FontSize = 10, VerticalAlignment = VerticalAlignment.Center });
            }
            var header = new Button { Style = (Style)Application.Current.Resources["PaneCommand"], Content = text, Padding = new Thickness(4, 6, 4, 6), Margin = new Thickness(0, 0, 4, 0) };
            ToolTipService.SetToolTip(header, $"Sort by {reportColumns[i]}");
            Microsoft.UI.Xaml.Automation.AutomationProperties.SetName(header, $"Sort by {reportColumns[i]}");
            var column = i;
            header.Click += (_, _) =>
            {
                reportDescending = reportSort == column && !reportDescending;
                reportSort = column;
                ApplyReport();
            };
            Grid.SetColumn(header, i);
            ReportHeader.Children.Add(header);
        }
    }

    /// <summary>A row as it shows: a cell per column, laid out as the header (recycled rows lay out again when the columns change).</summary>
    private void OnReportRowShown(ListViewBase sender, ContainerContentChangingEventArgs args)
    {
        if (args.InRecycleQueue || args.ItemContainer.ContentTemplateRoot is not Grid grid || args.Item is not ReportRow row)
        {
            return;
        }
        if (!Equals(grid.Tag, reportVersion))
        {
            grid.Children.Clear();
            grid.ColumnDefinitions.Clear();
            for (var i = 0; i < reportColumns.Count; i++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = reportWidths[i], MinWidth = 44 });
                var cell = new TextBlock
                {
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(8, 0, 8, 0),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    IsTextSelectionEnabled = false,
                };
                Grid.SetColumn(cell, i);
                grid.Children.Add(cell);
            }
            grid.Tag = reportVersion;
        }
        for (var i = 0; i < reportColumns.Count && i < grid.Children.Count; i++)
        {
            var cell = (TextBlock)grid.Children[i];
            var value = i < row.Cells.Count ? row.Cells[i] : "";
            cell.Text = value;
            ToolTipService.SetToolTip(cell, value.Length > 24 ? value : null);
        }
        args.Handled = true;
    }

    // ---- copying and saving ----

    private void OnReportRowContext(UIElement sender, ContextRequestedEventArgs e)
    {
        ListViewItem? container = null;
        for (var el = e.OriginalSource as DependencyObject; el is not null && el != ReportList; el = Microsoft.UI.Xaml.Media.VisualTreeHelper.GetParent(el))
        {
            if (el is ListViewItem item)
            {
                container = item;
                break;
            }
        }
        if (container is null || ReportList.ItemFromContainer(container) is not ReportRow row)
        {
            return;
        }
        e.Handled = true;
        var menu = new MenuFlyout();
        void Add(string text, string glyph, Action act)
        {
            var item = new MenuFlyoutItem { Text = text, Icon = new FontIcon { Glyph = glyph } };
            item.Click += (_, _) => act();
            menu.Items.Add(item);
        }
        Add("Copy the row", "", () => CopyText(string.Join("\t", row.Cells)));
        foreach (var address in row.Cells.Where(c => c.StartsWith("https://", StringComparison.OrdinalIgnoreCase)).Distinct().Take(2))
        {
            Add($"Open {address}", "", () => OpenUrl(address));
        }
        if (row.Target.Kind != NodeKind.Tenant)
        {
            Add($"Open {row.Target.Title} in the browser", "", () => OpenUrl(row.Target.Url));
        }
        menu.Items.Add(new MenuFlyoutSeparator());
        Add("Copy the table", "", CopyReport);
        if (e.TryGetPosition(container, out var at))
        {
            menu.ShowAt(container, at);
        }
        else
        {
            menu.ShowAt(container);
        }
    }

    private void OnCopyReport(object sender, RoutedEventArgs e) => CopyReport();

    /// <summary>The rows shown, as they're sorted, tab-separated with the header: pasted into Excel, they fill its cells.</summary>
    private void CopyReport()
    {
        var text = new StringBuilder();
        foreach (var cells in reportShown.Select(r => r.Cells).Prepend(reportColumns))
        {
            text.AppendLine(string.Join("\t", cells.Select(c => c.Replace('\t', ' ').ReplaceLineEndings(" "))));
        }
        CopyText(text.ToString());
        ShowToast("Copied the table", $"{Plural(reportShown.Count, "row", "rows")}, ready to paste into Excel.");
    }

    private async void OnSaveReportCsv(object sender, RoutedEventArgs e)
    {
        var name = string.Concat($"{reportTitle} {DateTime.Now:yyyy-MM-dd}".Select(ch => Path.GetInvalidFileNameChars().Contains(ch) ? '-' : ch));
        if (await PickSaveFileAsync(name, "CSV for Excel", ".csv") is not { } path)
        {
            return;
        }
        try
        {
            await File.WriteAllTextAsync(path, ResultsCsv.Table(reportColumns, reportShown.Select(r => r.Cells)), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            ShowError(new InvalidOperationException($"Couldn't save the report: {ex.Message}"));
            return;
        }
        ShowToast($"Saved {Path.GetFileName(path)}", $"{Plural(reportShown.Count, "row", "rows")}, as they're sorted and filtered.", ("Open", () =>
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true })?.Dispose();
            }
            catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException)
            {
                ShowError(new InvalidOperationException($"Couldn't open it: {ex.Message}"));
            }
        }));
    }
}
