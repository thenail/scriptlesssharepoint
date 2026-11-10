using System.Text.Json;
using System.Text.Json.Serialization;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.History;

/// <summary>How one target of a run went.</summary>
public enum RunOutcome
{
    /// <summary>Not over: still running, or the app closed while it ran.</summary>
    Pending,
    Done,
    /// <summary>Nothing needed doing.</summary>
    Unchanged,
    /// <summary>Asked to confirm it, the user skipped it.</summary>
    NotRun,
    Failed,
}

/// <summary>One target of a run: what it was, where it showed in the table, and how it went.</summary>
public sealed record RunTarget
{
    public required SpNode Target { get; init; }

    /// <summary>The site collection it's in, as the table had it.</summary>
    public SpNode? Site { get; init; }

    /// <summary>The subsite, list or folder it was ticked in.</summary>
    public SpNode? Parent { get; init; }

    public RunOutcome Outcome { get; init; }
    public string Message { get; init; } = "";
    public string? Detail { get; init; }

    public static RunOutcome OutcomeOf(CommandResult r) =>
        !r.Ok ? RunOutcome.Failed : r.WasDeclined ? RunOutcome.NotRun : r.WasSkipped ? RunOutcome.Unchanged : RunOutcome.Done;
}

/// <summary>
/// A command run from the app: what, where, with which inputs, and how each
/// target went. <see cref="File"/> is the file it's kept in.
/// </summary>
public sealed record RunRecord
{
    [JsonIgnore]
    public string File { get; init; } = "";

    public required string CommandId { get; init; }
    public required string Title { get; init; }
    public required string ConnectionId { get; init; }
    public string ConnectionName { get; init; } = "";

    /// <summary>The kind of target it ran on (<see cref="NodeKind.Tenant"/>: the tenant itself).</summary>
    public NodeKind Kind { get; init; }

    public DateTimeOffset Started { get; init; }

    /// <summary>When it ended; null while it runs, or when the app closed first.</summary>
    public DateTimeOffset? Finished { get; init; }

    public bool Cancelled { get; init; }

    /// <summary>The inputs given, as the command reads them.</summary>
    public IReadOnlyDictionary<string, string?> Args { get; init; } = new Dictionary<string, string?>();

    /// <summary>The inputs as shown ("Show in search: off").</summary>
    public IReadOnlyList<string> Inputs { get; init; } = [];

    public IReadOnlyList<RunTarget> Targets { get; init; } = [];

    public int Count(RunOutcome outcome) => Targets.Count(t => t.Outcome == outcome);
}

/// <summary>
/// The run history: one JSON file per run in <see cref="Folder"/>
/// (History\ in the data folder), the newest <see cref="Keep"/> kept.
/// </summary>
public sealed class RunHistoryStore(string? folder = null)
{
    /// <summary>How many runs are kept; older ones go as new ones come.</summary>
    public const int Keep = 500;

    /// <summary>The most of a target's detail kept (a read can bring a lot).</summary>
    public const int DetailLimit = 20_000;

    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = false,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() },
    };

    public string Folder { get; } = folder ?? Path.Combine(AppInfo.DataDir, "History");

    /// <summary>Saves a run (again, as it goes on), returning it with its file. The oldest beyond <see cref="Keep"/> go.</summary>
    public RunRecord Save(RunRecord run)
    {
        Directory.CreateDirectory(Folder);
        var file = run.File.Length > 0 ? run.File : $"{run.Started.UtcDateTime:yyyyMMdd-HHmmss-fff}-{Guid.NewGuid().ToString("N")[..8]}.json";
        var saved = run with
        {
            File = file,
            Targets = run.Targets.Select(t => t.Detail is { Length: > DetailLimit } d ? t with { Detail = d[..DetailLimit] + "\n…" } : t).ToList(),
        };
        var path = Path.Combine(Folder, file);
        var temp = path + ".tmp";
        System.IO.File.WriteAllText(temp, JsonSerializer.Serialize(saved, Json));
        System.IO.File.Move(temp, path, overwrite: true);
        if (run.File.Length == 0)
        {
            Prune();
        }
        return saved;
    }

    /// <summary>Every run, the newest first; files that can't be read are left out.</summary>
    public IReadOnlyList<RunRecord> List()
    {
        if (!Directory.Exists(Folder))
        {
            return [];
        }
        return Directory.EnumerateFiles(Folder, "*.json")
            .Select(f => Load(Path.GetFileName(f)))
            .OfType<RunRecord>()
            .OrderByDescending(r => r.Started)
            .ToList();
    }

    public RunRecord? Load(string file)
    {
        try
        {
            var path = Path.Combine(Folder, Path.GetFileName(file));
            return System.IO.File.Exists(path) && JsonSerializer.Deserialize<RunRecord>(System.IO.File.ReadAllText(path), Json) is { } r
                ? r with { File = Path.GetFileName(file) }
                : null;
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or JsonException or NotSupportedException)
        {
            return null;
        }
    }

    public void Delete(string file)
    {
        var path = Path.Combine(Folder, Path.GetFileName(file));
        if (System.IO.File.Exists(path))
        {
            System.IO.File.Delete(path);
        }
    }

    public void Clear()
    {
        if (Directory.Exists(Folder))
        {
            foreach (var f in Directory.EnumerateFiles(Folder, "*.json"))
            {
                System.IO.File.Delete(f);
            }
        }
    }

    /// <summary>Files are named by when the run started, so the oldest sort first.</summary>
    private void Prune()
    {
        var files = Directory.EnumerateFiles(Folder, "*.json").Order(StringComparer.Ordinal).ToList();
        foreach (var f in files.Take(Math.Max(0, files.Count - Keep)))
        {
            try
            {
                System.IO.File.Delete(f);
            }
            catch (IOException)
            {
            }
        }
    }
}
