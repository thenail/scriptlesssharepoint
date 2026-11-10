using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Plans;

/// <summary>
/// {Variables} in a plan's inputs, and their values for each item: the
/// built-in ones from the item itself, then the values given for everything,
/// then the item's row in a table (the most specific wins). A row is for the
/// item at its address and everything inside it, so libraries and files a
/// finder step found take their site's row.
/// </summary>
public static partial class PlanValues
{
    /// <summary>The built-in variables, from the item each step runs on.</summary>
    public static readonly IReadOnlyList<string> BuiltIn = ["title", "url", "name", "site", "web", "today"];

    [GeneratedRegex(@"\{([A-Za-z][A-Za-z0-9_]*)\}")]
    private static partial Regex Token();

    [GeneratedRegex(@"^[A-Za-z][A-Za-z0-9_]*$")]
    private static partial Regex Name();

    public static bool IsName(string name) => Name().IsMatch(name);

    /// <summary>The variables a value uses: "{title} (archived)" uses title.</summary>
    public static IEnumerable<string> Used(string? text) =>
        text is null ? [] : Token().Matches(text).Select(m => m.Groups[1].Value);

    /// <summary>A value with its variables filled in; ones without a value stay as they are.</summary>
    public static string Resolve(string text, IReadOnlyDictionary<string, string> values) =>
        Token().Replace(text, m => values.TryGetValue(m.Groups[1].Value, out var v) ? v : m.Value);

    /// <summary>The built-in values for an item.</summary>
    public static Dictionary<string, string> BuiltIns(SpNode n, DateTimeOffset now) => new(StringComparer.OrdinalIgnoreCase)
    {
        ["title"] = n.Title,
        ["url"] = n.Url,
        ["name"] = n.Kind is NodeKind.SiteCollection or NodeKind.Subsite or NodeKind.Tenant
            ? n.Url.TrimEnd('/')[(n.Url.TrimEnd('/').LastIndexOf('/') + 1)..]
            : n.LeafName,
        ["site"] = (n.SiteUrl ?? n.WebUrl).TrimEnd('/'),
        ["web"] = n.WebUrl.TrimEnd('/'),
        ["today"] = now.LocalDateTime.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
    };

    /// <summary>Whether a row's address is the item's, or a site or folder it's in.</summary>
    public static bool Covers(string rowUrl, string itemUrl)
    {
        var row = Uri.UnescapeDataString(rowUrl.Trim()).TrimEnd('/');
        var item = Uri.UnescapeDataString(itemUrl.Trim()).TrimEnd('/');
        return string.Equals(row, item, StringComparison.OrdinalIgnoreCase) || item.StartsWith(row + "/", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>The row for an item: its own, else the nearest one around it.</summary>
    public static PlanRow? RowFor(SpNode n, IReadOnlyList<PlanRow> rows) =>
        rows.Where(r => Covers(r.Url, n.Url)).MaxBy(r => r.Url.TrimEnd('/').Length);

    /// <summary>All the values for an item: built-in, then shared, then its row's.</summary>
    public static Dictionary<string, string> For(SpNode n, IReadOnlyDictionary<string, string> shared, IReadOnlyList<PlanRow> rows, DateTimeOffset now)
    {
        var values = BuiltIns(n, now);
        foreach (var (k, v) in shared.Where(x => !string.IsNullOrWhiteSpace(x.Value)))
        {
            values[k] = v;
        }
        if (RowFor(n, rows) is { } row)
        {
            foreach (var (k, v) in row.Values.Where(x => !string.IsNullOrWhiteSpace(x.Value)))
            {
                values[k] = v;
            }
        }
        return values;
    }

    /// <summary>A step's inputs for one item: its defaults, then the step's, with the item's values filled in.</summary>
    public static CommandArgs ArgsFor(Command command, PlanStep step, IReadOnlyDictionary<string, string> values)
    {
        var args = PlanCheck.ArgsOf(command, step);
        foreach (var (name, value) in args.ToDictionary().ToList())
        {
            if (value is not null && Token().IsMatch(value))
            {
                args[name] = Resolve(value, values);
            }
        }
        return args;
    }

    /// <summary>
    /// Inputs for checking a plan before its values are known: those with
    /// {variables} get a value that passes, so only the rest are checked.
    /// </summary>
    public static CommandArgs WithSamples(Command command, CommandArgs args)
    {
        var checkable = new CommandArgs();
        foreach (var p in command.Parameters)
        {
            var value = args[p.Name];
            checkable[p.Name] = value is not null && Token().IsMatch(value) ? Sample(p) : value;
        }
        return checkable;
    }

    private static string Sample(CommandParameter p) => p.Kind switch
    {
        ParameterKind.Number => (p.Min ?? 1).ToString(CultureInfo.InvariantCulture),
        ParameterKind.Toggle => "true",
        ParameterKind.Choice => p.Choices.FirstOrDefault()?.Value ?? "x",
        ParameterKind.Url => "https://contoso.sharepoint.com/sites/x",
        ParameterKind.Person => "someone@contoso.com",
        _ => "x",
    };

    /// <summary>
    /// A table as pasted from a spreadsheet or read from a CSV file: a header
    /// row (the address column first, then one column per variable), then a
    /// row per item. Tabs, semicolons or commas separate columns; quotes work
    /// as in CSV. Returns the rows and what's wrong with it.
    /// </summary>
    public static (IReadOnlyList<PlanRow> Rows, IReadOnlyList<string> Columns, IReadOnlyList<string> Problems) ParseTable(string text)
    {
        var lines = text.Replace("\r\n", "\n").Split('\n').Where(l => l.Trim().Length > 0).ToList();
        if (lines.Count == 0)
        {
            return ([], [], []);
        }
        var header = lines[0];
        var separator = header.Contains('\t') ? '\t' : header.Count(c => c == ';') > header.Count(c => c == ',') ? ';' : ',';
        var columns = Fields(header, separator).Select(c => c.Trim().Trim('{', '}')).ToList();
        var problems = new List<string>();
        if (columns.Count < 2)
        {
            problems.Add("The first row names the columns: the address, then one per variable (separated by tabs, commas or semicolons).");
            return ([], columns, problems);
        }
        var rows = new List<PlanRow>();
        for (var i = 1; i < lines.Count; i++)
        {
            var fields = Fields(lines[i], separator);
            var url = fields.ElementAtOrDefault(0)?.Trim() ?? "";
            if (url.Length == 0)
            {
                continue;
            }
            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeHttps)
            {
                problems.Add($"Row {i + 1}: \"{url}\" isn't an https address.");
                continue;
            }
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (var c = 1; c < columns.Count; c++)
            {
                values[columns[c]] = fields.ElementAtOrDefault(c)?.Trim() ?? "";
            }
            rows.Add(new PlanRow(url.TrimEnd('/'), values));
        }
        foreach (var dup in rows.GroupBy(r => r.Url, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
        {
            problems.Add($"{dup.Key} has more than one row.");
        }
        return (rows, columns.Skip(1).ToList(), problems);
    }

    /// <summary>A table as text again (tab-separated), to show and edit.</summary>
    public static string TableText(IReadOnlyList<PlanRow> rows, IEnumerable<string> columns)
    {
        var cols = columns.ToList();
        var text = new StringBuilder("URL");
        foreach (var c in cols)
        {
            text.Append('\t').Append(c);
        }
        foreach (var r in rows)
        {
            text.Append('\n').Append(r.Url);
            foreach (var c in cols)
            {
                text.Append('\t').Append(r.Values.GetValueOrDefault(c, ""));
            }
        }
        return text.ToString();
    }

    private static List<string> Fields(string line, char separator)
    {
        var fields = new List<string>();
        var field = new StringBuilder();
        var quoted = false;
        for (var i = 0; i < line.Length; i++)
        {
            var ch = line[i];
            if (quoted)
            {
                if (ch == '"' && i + 1 < line.Length && line[i + 1] == '"')
                {
                    field.Append('"');
                    i++;
                }
                else if (ch == '"')
                {
                    quoted = false;
                }
                else
                {
                    field.Append(ch);
                }
            }
            else if (ch == '"' && field.ToString().Trim().Length == 0)
            {
                field.Clear();
                quoted = true;
            }
            else if (ch == separator)
            {
                fields.Add(field.ToString());
                field.Clear();
            }
            else
            {
                field.Append(ch);
            }
        }
        fields.Add(field.ToString());
        return fields;
    }
}
