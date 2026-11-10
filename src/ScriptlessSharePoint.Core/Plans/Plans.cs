using System.Text.Json;
using System.Text.Json.Serialization;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Plans;

/// <summary>One step of a plan: a command from the catalog, with its inputs ({variables} allowed).</summary>
public sealed record PlanStep
{
    public required string Command { get; init; }
    public Dictionary<string, string?> Args { get; init; } = [];
}

/// <summary>A value a plan asks for when it runs: one for everything, or one per item from a table.</summary>
public sealed record PlanVariable
{
    /// <summary>As written in inputs: <c>{owner}</c>.</summary>
    public required string Name { get; init; }

    public string? Label { get; init; }

    /// <summary>Text, Person, Number, Url or LongText: how the app asks for it.</summary>
    public ParameterKind Kind { get; init; } = ParameterKind.Text;

    public string? Hint { get; init; }

    [JsonIgnore]
    public string Shown => string.IsNullOrWhiteSpace(Label) ? Name : Label;
}

/// <summary>One row of values: for the item at <see cref="Url"/>, and anything inside it.</summary>
public sealed record PlanRow(string Url, IReadOnlyDictionary<string, string> Values);

/// <summary>
/// A run made ready (by an AI app, or kept from the app): what to run on, and
/// the values: <see cref="Values"/> for everything, <see cref="Rows"/> per item.
/// </summary>
public sealed record PreparedRun
{
    public string? Preset { get; init; }
    public IReadOnlyList<SpNode> Items { get; init; } = [];
    public IReadOnlyDictionary<string, string> Values { get; init; } = new Dictionary<string, string>();
    public IReadOnlyList<PlanRow> Rows { get; init; } = [];
    public string? By { get; init; }
    public DateTimeOffset At { get; init; }
}

/// <summary>
/// Steps to run, in order, on whatever is ticked when it runs: a recipe, not
/// tied to items. Made by an AI app (through the MCP server) or in the app,
/// looked over and run in the app. Inputs can hold {variables}: built-in ones
/// (<see cref="PlanValues.BuiltIn"/>: the item's title, address…) and the
/// plan's own (<see cref="Variables"/>), given when it runs, once for all or
/// per item. A finder step ("Tick inside") replaces what the steps after it
/// run on with what it found; found items take the values of the row they're in.
/// </summary>
public sealed record Plan
{
    /// <summary>The file it's kept in, whose name (without .json) is its ID.</summary>
    [JsonIgnore]
    public string File { get; init; } = "";

    [JsonIgnore]
    public string Id => Path.GetFileNameWithoutExtension(File);

    public required string Name { get; init; }
    public required string ConnectionId { get; init; }
    public string? Description { get; init; }
    public DateTimeOffset Created { get; init; }
    public DateTimeOffset Changed { get; init; }

    /// <summary>Who made it: the AI app's name, or "you".</summary>
    public string? CreatedBy { get; init; }

    /// <summary>What it starts on (site collections, libraries…); empty: what its first step runs on.</summary>
    public IReadOnlyList<NodeKind> StartsOn { get; init; } = [];

    public IReadOnlyList<PlanVariable> Variables { get; init; } = [];

    public IReadOnlyList<PlanStep> Steps { get; init; } = [];

    // ---- a test selection: remembered, to try it on ----

    /// <summary>A predefined selection to try it on.</summary>
    public string? Preset { get; init; }

    /// <summary>Items to try it on.</summary>
    public IReadOnlyList<SpNode> Items { get; init; } = [];

    public IReadOnlyList<SpNode> Sites { get; init; } = [];
    public IReadOnlyList<SpNode> Parents { get; init; } = [];

    [JsonIgnore]
    public bool HasTestSelection => Preset is not null || Items.Count > 0;

    /// <summary>A run made ready to look over and start.</summary>
    public PreparedRun? Prepared { get; init; }
}

/// <summary>The plans: one JSON file each in <see cref="Folder"/> (Plans\ in the data folder), shared by the app and the MCP server.</summary>
public sealed class PlanStore(string? folder = null)
{
    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() },
    };

    public string Folder { get; } = folder ?? Path.Combine(AppInfo.DataDir, "Plans");

    /// <summary>Every plan, the newest first; files that can't be read are left out.</summary>
    public IReadOnlyList<Plan> List() => !Directory.Exists(Folder)
        ? []
        : Directory.EnumerateFiles(Folder, "*.json").Select(f => Load(Path.GetFileNameWithoutExtension(f))).OfType<Plan>()
            .OrderByDescending(p => p.Created).ToList();

    /// <summary>A plan by its ID (or file name); null when there's none.</summary>
    public Plan? Load(string id)
    {
        var path = PathOf(id);
        try
        {
            return System.IO.File.Exists(path) && JsonSerializer.Deserialize<Plan>(System.IO.File.ReadAllText(path), Json) is { } p
                ? p with { File = Path.GetFileName(path) }
                : null;
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or JsonException or NotSupportedException)
        {
            return null;
        }
    }

    /// <summary>Saves a plan: a new one (no file yet) gets a file of its own.</summary>
    public Plan Save(Plan plan)
    {
        Directory.CreateDirectory(Folder);
        var now = DateTimeOffset.Now;
        var file = plan.File.Length > 0 ? plan.File : $"{now:yyyyMMdd-HHmmss}-{Guid.NewGuid().ToString("N")[..8]}.json";
        var saved = plan with
        {
            File = file,
            Created = plan.File.Length > 0 && plan.Created != default ? plan.Created : now,
            Changed = now,
            Items = plan.Items.DistinctBy(n => n.Key).ToList(),
            Sites = plan.Sites.DistinctBy(n => n.Key).ToList(),
            Parents = plan.Parents.DistinctBy(n => n.Key).ToList(),
        };
        var path = PathOf(file);
        var temp = path + ".tmp";
        System.IO.File.WriteAllText(temp, JsonSerializer.Serialize(saved, Json));
        System.IO.File.Move(temp, path, overwrite: true);
        return saved;
    }

    public void Delete(string id)
    {
        var path = PathOf(id);
        if (System.IO.File.Exists(path))
        {
            System.IO.File.Delete(path);
        }
    }

    private string PathOf(string id)
    {
        var name = Path.GetFileName(id.Trim());
        return Path.Combine(Folder, name.EndsWith(".json", StringComparison.OrdinalIgnoreCase) ? name : name + ".json");
    }
}

/// <summary>
/// Checking a plan: each step's command exists, its inputs are right (inputs
/// with {variables} are checked when the values are known), it runs on what
/// the steps before it leave ticked, and its variables are declared.
/// </summary>
public static class PlanCheck
{
    /// <summary>The kinds a plan starts on: as it says, else its first step's (the tenant only when that's all it runs on).</summary>
    public static IReadOnlySet<NodeKind> StartKinds(Plan plan)
    {
        if (plan.StartsOn.Count > 0)
        {
            return plan.StartsOn.ToHashSet();
        }
        if (plan.Steps.Count > 0 && CommandCatalog.Find(plan.Steps[0].Command) is { } first)
        {
            var kinds = first.Targets.Where(k => k != NodeKind.Tenant).ToHashSet();
            return kinds.Count > 0 ? kinds : [];
        }
        return new HashSet<NodeKind>();
    }

    /// <summary>A step's inputs as the command reads them: its defaults, then what the step gives.</summary>
    public static CommandArgs ArgsOf(Command command, PlanStep step)
    {
        var args = command.Defaults();
        foreach (var (name, value) in step.Args)
        {
            args[name] = value;
        }
        return args;
    }

    /// <summary>What's wrong with a plan, one line each; empty when it can run.</summary>
    public static IReadOnlyList<string> Problems(Plan plan)
    {
        var problems = new List<string>();
        if (string.IsNullOrWhiteSpace(plan.Name))
        {
            problems.Add("It needs a name.");
        }
        if (plan.Preset is { } preset && SitePresets.Find(preset) is null && ContentPresets.Find(preset) is null)
        {
            problems.Add($"There's no predefined selection \"{preset}\".");
        }
        if (plan.Steps.Count == 0)
        {
            problems.Add("It has no steps.");
        }
        foreach (var v in plan.Variables)
        {
            if (!PlanValues.IsName(v.Name))
            {
                problems.Add($"Variable \"{v.Name}\": a name is letters, digits and _, starting with a letter.");
            }
            else if (PlanValues.BuiltIn.Contains(v.Name, StringComparer.OrdinalIgnoreCase))
            {
                problems.Add($"Variable \"{v.Name}\": that's a built-in one ({string.Join(", ", PlanValues.BuiltIn.Select(b => $"{{{b}}}"))}).");
            }
            if (v.Kind is not (ParameterKind.Text or ParameterKind.LongText or ParameterKind.Person or ParameterKind.Number or ParameterKind.Url))
            {
                problems.Add($"Variable \"{v.Name}\": its type is text, longText, person, number or url.");
            }
        }
        foreach (var dup in plan.Variables.GroupBy(v => v.Name, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
        {
            problems.Add($"Variable \"{dup.Key}\" is declared more than once.");
        }
        var known = PlanValues.BuiltIn.Concat(plan.Variables.Select(v => v.Name)).ToHashSet(StringComparer.OrdinalIgnoreCase);

        var kinds = StartKinds(plan);
        for (var i = 0; i < plan.Steps.Count; i++)
        {
            var step = plan.Steps[i];
            var where = $"Step {i + 1} ({step.Command})";
            if (CommandCatalog.Find(step.Command) is not { } command)
            {
                problems.Add($"{where}: there's no such command.");
                continue;
            }
            var unknown = step.Args.Keys.Where(k => command.Parameters.All(p => p.Name != k)).ToList();
            if (unknown.Count > 0)
            {
                problems.Add($"{where}: it has no input {string.Join(", ", unknown)}.");
            }
            foreach (var name in step.Args.Values.SelectMany(PlanValues.Used).Distinct(StringComparer.OrdinalIgnoreCase).Where(n => !known.Contains(n)))
            {
                problems.Add($"{where}: it uses {{{name}}}, which isn't one of the plan's variables.");
            }
            var on = RunsOn(command, kinds);
            if (on.Count == 0)
            {
                problems.Add($"{where}: it runs on {string.Join(", ", command.Targets.Select(k => NodeKinds.Name(k, 2)))}, " +
                             $"but what's ticked then is {(kinds.Count == 0 ? "nothing" : string.Join(", ", kinds.Select(k => NodeKinds.Name(k, 2))))}.");
            }
            else
            {
                problems.AddRange(command.Check(PlanValues.WithSamples(command, ArgsOf(command, step)), on).Select(p => $"{where}: {p}"));
            }
            kinds = KindsAfter(command, ArgsOf(command, step), kinds);
        }
        return problems;
    }

    /// <summary>The kinds a command runs on, of those ticked: the tenant when none fits and it runs on the tenant.</summary>
    public static IReadOnlyList<NodeKind> RunsOn(Command command, IReadOnlySet<NodeKind> kinds)
    {
        var on = kinds.Where(command.AppliesTo).ToList();
        return on.Count == 0 && command.AppliesTo(NodeKind.Tenant) ? [NodeKind.Tenant] : on;
    }

    /// <summary>Whether a plan can start on what's ticked: its first step runs on some of it (or on the tenant).</summary>
    public static bool StartsOnAny(Plan plan, IEnumerable<NodeKind> ticked) =>
        plan.Steps.Count > 0 && CommandCatalog.Find(plan.Steps[0].Command) is { } first && RunsOn(first, ticked.ToHashSet()).Count > 0;

    /// <summary>What's ticked after a step: a finder's finds, else as it was.</summary>
    public static IReadOnlySet<NodeKind> KindsAfter(Command command, CommandArgs args, IReadOnlySet<NodeKind> kinds) =>
        command.After != After.Tick ? kinds : command.Id switch
        {
            "find.subsites" => new HashSet<NodeKind> { NodeKind.Subsite },
            "find.lists" or "find.list-state" => new HashSet<NodeKind> { NodeKind.List, NodeKind.Library },
            "find.home" => new HashSet<NodeKind> { NodeKind.File },
            "find.content" => args["what"] switch
            {
                "files" => new HashSet<NodeKind> { NodeKind.File },
                "folders" => new HashSet<NodeKind> { NodeKind.Folder },
                "items" => new HashSet<NodeKind> { NodeKind.Item },
                _ => new HashSet<NodeKind> { NodeKind.File, NodeKind.Folder, NodeKind.Item },
            },
            "find.search" => new HashSet<NodeKind> { NodeKind.File, NodeKind.Item, NodeKind.List, NodeKind.Library },
            _ => new HashSet<NodeKind> { NodeKind.List, NodeKind.Library, NodeKind.Folder, NodeKind.File, NodeKind.Item },
        };

    /// <summary>
    /// Before a run: what's wrong with running the plan on these items with
    /// these values. Each of the plan's variables a step uses needs a value
    /// for every item; the steps before any finder have their inputs checked
    /// item by item (later ones run on what's found, so they're checked as they run).
    /// Notes are for rows that match nothing ticked.
    /// </summary>
    public static (IReadOnlyList<string> Problems, IReadOnlyList<string> Notes) RunProblems(
        Plan plan, IReadOnlyList<SpNode> items, IReadOnlyDictionary<string, string> shared, IReadOnlyList<PlanRow> rows, DateTimeOffset now)
    {
        var problems = new List<string>();
        var used = plan.Steps.SelectMany(s => s.Args.Values.SelectMany(PlanValues.Used)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var v in plan.Variables.Where(v => used.Contains(v.Name)))
        {
            var missing = items.Where(n => !PlanValues.For(n, shared, rows, now).TryGetValue(v.Name, out var x) || string.IsNullOrWhiteSpace(x)).ToList();
            if (missing.Count > 0)
            {
                problems.Add($"{v.Shown}: no value for {Some(missing)}.");
            }
        }
        var kinds = items.Select(n => n.Kind).ToHashSet();
        foreach (var step in plan.Steps)
        {
            if (CommandCatalog.Find(step.Command) is not { } command)
            {
                break;
            }
            var on = RunsOn(command, kinds);
            var targets = on.Contains(NodeKind.Tenant) ? [] : items.Where(n => on.Contains(n.Kind)).ToList();
            var wrong = targets
                .SelectMany(n => command.Check(PlanValues.ArgsFor(command, step, PlanValues.For(n, shared, rows, now)), [n.Kind]).Select(p => (p, n)))
                .GroupBy(x => x.p)
                .Select(g => $"{command.Title}: {g.Key.TrimEnd('.')} for {Some(g.Select(x => x.n).ToList())}.");
            problems.AddRange(wrong.Take(10));
            if (command.After == After.Tick)
            {
                break;
            }
        }
        var notes = new List<string>();
        var unmatched = rows.Where(r => !items.Any(n => PlanValues.Covers(r.Url, n.Url))).ToList();
        if (unmatched.Count > 0)
        {
            notes.Add($"{unmatched.Count:N0} of the table's rows match nothing ticked ({string.Join(", ", unmatched.Take(3).Select(r => r.Url))}{(unmatched.Count > 3 ? "…" : "")}).");
        }
        return (problems, notes);

        static string Some(IReadOnlyList<SpNode> nodes) =>
            nodes.Count == 1 ? nodes[0].Title : $"{nodes.Count:N0} of them ({string.Join(", ", nodes.Take(3).Select(n => n.Title))}{(nodes.Count > 3 ? "…" : "")})";
    }
}
