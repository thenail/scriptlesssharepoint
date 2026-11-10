using System.Globalization;
using System.Text;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.History;

/// <summary>One target's result, for a CSV: the step (a plan's), what it ran on, how it went, what it said, and what it listed.</summary>
public sealed record ResultLine(int? Step, SpNode Target, RunOutcome Outcome, string Message, string? Detail, ReportTable? Table = null);

/// <summary>
/// Results as CSV, for Excel: a row per thing a report lists (List users, List
/// columns…) under its columns' names (or the "  ·  " parts of the lines it
/// listed, as "Detail 1"…), else a row per target. The step column is there
/// only for a plan's results.
/// </summary>
public static class ResultsCsv
{
    public static string Build(IEnumerable<ResultLine> lines)
    {
        // each one's table, else the lines it listed, split into their parts; their columns together, by name
        var items = lines.Select(l => (Line: l, Table: l.Table ?? (string.IsNullOrWhiteSpace(l.Detail) ? null : ReportTable.FromLines(l.Detail)))).ToList();
        var steps = items.Any(x => x.Line.Step is not null);
        var columns = ReportTable.Union(items.Select(x => x.Table).OfType<ReportTable>());
        var rows = new List<string[]>();
        rows.Add(["Step", "Name", "Address", "Result", "Message", .. columns]);
        foreach (var (l, table) in items)
        {
            string[] head = [l.Step?.ToString(CultureInfo.InvariantCulture) ?? "", l.Target.Title, l.Target.Url, OutcomeText(l.Outcome), l.Message];
            if (table is not { Rows.Count: > 0 })
            {
                rows.Add([.. head, .. columns.Select(_ => "")]);
                continue;
            }
            rows.AddRange(table.Rows.Select(r => (string[])[.. head, .. table.Cells(r, columns)]));
        }
        var csv = new StringBuilder();
        foreach (var row in rows)
        {
            csv.Append(string.Join(",", row.Skip(steps ? 0 : 1).Select(Cell))).Append("\r\n");
        }
        return csv.ToString();
    }

    /// <summary>A table as it's shown (a report, sorted and filtered): its columns, then its rows.</summary>
    public static string Table(IReadOnlyList<string> columns, IEnumerable<IReadOnlyList<string>> rows)
    {
        var csv = new StringBuilder();
        foreach (var row in rows.Prepend(columns))
        {
            csv.Append(string.Join(",", row.Select(Cell))).Append("\r\n");
        }
        return csv.ToString();
    }

    public static string OutcomeText(RunOutcome outcome) => outcome switch
    {
        RunOutcome.Done => "Done",
        RunOutcome.Unchanged => "Unchanged",
        RunOutcome.NotRun => "Not run",
        RunOutcome.Failed => "Failed",
        _ => "Didn't finish",
    };

    /// <summary>
    /// A cell, quoted when it has to be. Text from SharePoint that starts like
    /// a formula gets a ' first, so Excel shows it rather than running it.
    /// </summary>
    private static string Cell(string value)
    {
        if (value.Length > 0 && value[0] is '=' or '+' or '-' or '@' or '\t' or '\r')
        {
            value = "'" + value;
        }
        return value.IndexOfAny([',', '"', '\n', '\r']) >= 0 ? $"\"{value.Replace("\"", "\"\"")}\"" : value;
    }
}
