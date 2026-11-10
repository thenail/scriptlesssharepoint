// Plans: the commands and predefined selections an AI app can build a plan
// from; making, reading, changing and deleting plans; and preparing a run (what
// to run on, and per-item values). A plan changes nothing in SharePoint: the
// user runs it in the app. Its results come back from the app's run history.
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.History;
using ScriptlessSharePoint.Plans;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Mcp;

public sealed partial class SharePointTools
{
    /// <summary>What can be ticked, as tools name it (a property: the tool list is built from it, in another file's static field).</summary>
    private static string[] Kinds => ["tenant", "siteCollection", "subsite", "list", "library", "folder", "file", "item"];

    private static string[] VariableTypes => ["text", "longText", "person", "number", "url"];

    private static IEnumerable<McpTool> PlanTools()
    {
        var local = new ToolAnnotations(false, false, false, false);

        JsonObject Steps() => new()
        {
            ["type"] = "array",
            ["minItems"] = 1,
            ["description"] = "The steps, in order. Each runs on what's ticked: what the plan is run on, or what the last finder step found. " +
                              "Inputs can use {variables}: built-in ones ({title}, {url}, {name}, {site}, {web}, {today}: the item's own) and the plan's `variables`.",
            ["items"] = new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["command"] = Str("A command's id, from list_commands."),
                    ["args"] = new JsonObject
                    {
                        ["type"] = "object",
                        ["description"] = "Its inputs by parameter name; left out ones take their defaults. Choices by their value; toggles \"true\" or \"false\"; \"{owner}\" for a variable.",
                        ["additionalProperties"] = new JsonObject { ["type"] = new JsonArray("string", "number", "boolean") },
                    },
                },
                ["required"] = new JsonArray("command"),
                ["additionalProperties"] = false,
            },
        };
        JsonObject Variables() => new()
        {
            ["type"] = "array",
            ["description"] = "Values the plan asks for when it runs, used in inputs as {name}: one value for everything, or one per item from a table (prepare_run).",
            ["items"] = new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["name"] = Str("Letters, digits and _, starting with a letter: {name} in inputs."),
                    ["label"] = Str("How the app asks for it."),
                    ["type"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray([.. VariableTypes.Select(t => (JsonNode)t)]), ["default"] = "text" },
                    ["hint"] = Str("A hint shown with it."),
                },
                ["required"] = new JsonArray("name"),
                ["additionalProperties"] = false,
            },
        };
        JsonObject Sites(string what) => new() { ["type"] = "array", ["items"] = Str("A site collection's URL."), ["description"] = what };
        JsonObject Kind(string what) => new() { ["type"] = "string", ["enum"] = new JsonArray([.. Kinds.Select(k => (JsonNode)k)]), ["description"] = what };

        yield return new("list_commands", "List commands",
            "The commands a plan's steps can use: id, what they run on (`runsOn`), whether they're finders (they tick what they find for the next steps), whether they change or delete things, and their parameters with types, defaults and choices. Filter by `kind` (what's ticked) and `query` (words in the title, description or PowerShell cmdlet).",
            Schema(("kind", Kind("What's ticked."), false), ("query", Str("Words to look for."), false)),
            Read with { OpenWorldHint = false });
        yield return new("list_presets", "List predefined selections",
            "The predefined selections a plan can be run on: site collections or OneDrives picked by a rule from the admin centre (inactive, hub sites, sharing…), or files and pages found by search (large, old, guests can open…).",
            Schema(), Read with { OpenWorldHint = false });
        yield return new("list_plans", "List plans",
            "The plans saved in the app (for every connection, or one): id, name, what they start on, their variables and steps, and who made them.",
            Schema(Connection()), Read with { OpenWorldHint = false });
        yield return new("get_plan", "Get a plan",
            "A plan in full: what it starts on, its variables, its steps with their inputs, its test selection, the run prepared for it, anything wrong with it, and how each step went the last time the user ran it.",
            Schema(("id", Str("The plan's id."), true)), Read with { OpenWorldHint = false });
        yield return new("create_plan", "Create a plan",
            "Saves a plan: steps, and the variables they use. A plan isn't tied to items: it runs on whatever the user ticks, or what prepare_run gives it. Every step is checked (the command exists, its inputs are right, it runs on what's ticked then, its variables are declared); a plan with problems isn't saved, and the answer says what to fix. It changes nothing in SharePoint.",
            Schema(
                ("name", Str("A short name, as the app lists it."), true),
                ("description", Str("What it's for, for the user."), false),
                ("startsOn", Kind("What it runs on first (default: what its first step runs on)."), false),
                ("variables", Variables(), false),
                ("steps", Steps(), true),
                ("sites", Sites("Site collections to try it on (its test selection), from search_sites. Optional."), false),
                ("preset", Str("A predefined selection to try it on (from list_presets). Optional."), false),
                Connection()),
            local);
        yield return new("update_plan", "Change a plan",
            "Changes a plan: what's given replaces what it had (variables and steps as a whole). Checked as create_plan checks.",
            Schema(
                ("id", Str("The plan's id."), true),
                ("name", Str("A new name."), false),
                ("description", Str("A new description."), false),
                ("startsOn", Kind("What it runs on first."), false),
                ("variables", Variables(), false),
                ("steps", Steps(), false),
                ("sites", Sites("A new test selection."), false),
                ("preset", Str("A predefined selection as its test selection."), false)),
            local with { IdempotentHint = true });
        yield return new("prepare_run", "Prepare a run of a plan",
            "Makes a run of a plan ready in the app: what to run it on (`sites`, `preset`, or the URLs in `rows`) and its variables' values: `values` for everything, `rows` per item (a row is for that URL and everything inside it, so a site's row also covers its libraries and files). It's checked first: every item needs a value for each variable the steps use, and values must fit their inputs. Nothing runs: the user opens the plan, looks it over, and runs it.",
            Schema(
                ("id", Str("The plan's id."), true),
                ("sites", Sites("The site collections to run it on. Left out: the URLs in `rows`."), false),
                ("preset", Str("A predefined selection to run it on, instead of sites."), false),
                ("values", new JsonObject
                {
                    ["type"] = "object",
                    ["description"] = "Values for every item, by variable name.",
                    ["additionalProperties"] = new JsonObject { ["type"] = new JsonArray("string", "number", "boolean") },
                }, false),
                ("rows", new JsonObject
                {
                    ["type"] = "array",
                    ["description"] = "Values per item: a URL and its values by variable name.",
                    ["items"] = new JsonObject
                    {
                        ["type"] = "object",
                        ["properties"] = new JsonObject
                        {
                            ["url"] = Str("A site collection's (or other item's) URL."),
                            ["values"] = new JsonObject
                            {
                                ["type"] = "object",
                                ["additionalProperties"] = new JsonObject { ["type"] = new JsonArray("string", "number", "boolean") },
                            },
                        },
                        ["required"] = new JsonArray("url", "values"),
                        ["additionalProperties"] = false,
                    },
                }, false)),
            local with { IdempotentHint = true });
        yield return new("delete_plan", "Delete a plan",
            "Deletes a plan from the app. Nothing in SharePoint changes. Only when the user asks for it.",
            Schema(("id", Str("The plan's id."), true)),
            local with { DestructiveHint = true, IdempotentHint = true });
    }

    /// <summary>A plan tool's answer, or null when <paramref name="name"/> isn't one.</summary>
    private async Task<ToolResult?> RunPlanToolAsync(string name, Args a)
    {
        switch (name)
        {
            case "list_commands":
            {
                var kind = KindOf(a, "kind", name);
                var words = a.String("query")?.Split(' ', StringSplitOptions.RemoveEmptyEntries) ?? [];
                var commands = CommandCatalog.All
                    .Where(c => kind is null || c.AppliesTo(kind.Value))
                    .Where(c => words.All(w => c.Title.Contains(w, StringComparison.OrdinalIgnoreCase) ||
                                              c.Description.Contains(w, StringComparison.OrdinalIgnoreCase) ||
                                              c.PowerShell?.Contains(w, StringComparison.OrdinalIgnoreCase) == true))
                    .Where(c => c.Id != "open")
                    .Select(CommandView)
                    .ToList();
                return Data(new { commands, total = commands.Count }, "commands");
            }
            case "list_presets":
            {
                var presets = SitePresets.All
                    .Select(p => new { p.Id, p.Title, p.Description, picks = p.OneDrive ? "oneDrives" : "siteCollections" })
                    .Concat(ContentPresets.All.Select(p => new { p.Id, p.Title, p.Description, picks = "filesAndPages" }))
                    .ToList();
                return Data(new { presets }, "presets");
            }
            case "list_plans":
            {
                var connections = backend.ListConnections().ToDictionary(c => c.Id, c => c.Name);
                var only = a.String("connection") is { } c ? backend.Resolve(c).Id : null;
                var plans = backend.Plans.List()
                    .Where(p => only is null || p.ConnectionId == only)
                    .Select(p => new
                    {
                        p.Id,
                        p.Name,
                        p.Description,
                        connection = connections.GetValueOrDefault(p.ConnectionId, p.ConnectionId),
                        startsOn = PlanCheck.StartKinds(p),
                        variables = p.Variables.Select(v => v.Name).ToList(),
                        steps = p.Steps.Select(s => s.Command).ToList(),
                        prepared = p.Prepared is { } r ? Selection(r.Preset, r.Items, r.Rows) : null,
                        p.CreatedBy,
                        p.Changed,
                    })
                    .ToList();
                return Data(new { plans, total = plans.Count }, "plans");
            }
            case "get_plan":
            {
                var plan = Plan(a.Required("id"));
                var runs = backend.History.List().Where(r => r.Plan == plan.Id).ToList();
                return Data(new
                {
                    plan.Id,
                    plan.Name,
                    plan.Description,
                    connection = backend.ListConnections().FirstOrDefault(c => c.Id == plan.ConnectionId)?.Name ?? plan.ConnectionId,
                    plan.CreatedBy,
                    plan.Created,
                    plan.Changed,
                    startsOn = PlanCheck.StartKinds(plan),
                    variables = plan.Variables.Select(v => new { v.Name, v.Label, type = v.Kind, v.Hint }).ToList(),
                    steps = plan.Steps.Select((s, i) => new
                    {
                        step = i + 1,
                        s.Command,
                        title = CommandCatalog.Find(s.Command)?.Title,
                        args = s.Args,
                        lastRun = runs.Where(r => r.Step == i + 1).OrderByDescending(r => r.Started).Select(RunView).FirstOrDefault(),
                    }).ToList(),
                    testSelection = plan.HasTestSelection ? Selection(plan.Preset, plan.Items, []) : null,
                    prepared = plan.Prepared is { } r
                        ? new
                        {
                            runsOn = Selection(r.Preset, r.Items, r.Rows),
                            values = r.Values,
                            rows = r.Rows.Count,
                            r.By,
                            r.At,
                        }
                        : null,
                    problems = PlanCheck.Problems(plan) is { Count: > 0 } p ? p : null,
                }, "plan");
            }
            case "create_plan":
            {
                var connection = backend.Resolve(a.String("connection"));
                var plan = new Plan
                {
                    Name = a.Required("name"),
                    Description = a.String("description"),
                    ConnectionId = connection.Id,
                    CreatedBy = Client,
                    StartsOn = KindOf(a, "startsOn", name) is { } k ? [k] : [],
                    Variables = VariablesOf(a, name) ?? [],
                    Steps = StepsOf(a) ?? throw new ToolArgumentException($"Invalid arguments for {name}: steps: required"),
                };
                plan = WithTestSelection(plan, a, name);
                return Saved(plan, $"Saved. It runs on what the user ticks in {AppInfo.Name}; prepare_run can make a run ready (what to run on, and per-item values). open_app with `plan` shows it.");
            }
            case "update_plan":
            {
                var plan = Plan(a.Required("id"));
                plan = plan with
                {
                    Name = a.String("name") ?? plan.Name,
                    Description = a.String("description") ?? plan.Description,
                    StartsOn = KindOf(a, "startsOn", name) is { } k ? [k] : plan.StartsOn,
                    Variables = VariablesOf(a, name) ?? plan.Variables,
                    Steps = StepsOf(a) ?? plan.Steps,
                };
                plan = WithTestSelection(plan, a, name);
                return Saved(plan, "Changed.");
            }
            case "prepare_run":
            {
                var plan = Plan(a.Required("id"));
                if (PlanCheck.Problems(plan) is { Count: > 0 } planProblems)
                {
                    return new ToolResult("The plan has problems to fix first (update_plan):\n- " + string.Join("\n- ", planProblems), true);
                }
                var preset = a.String("preset");
                var rows = RowsOf(a, name);
                var values = ValuesOf(a.Node("values") as JsonObject);
                var sites = SiteNodes(a, "sites", name);
                if (preset is not null && sites is not null)
                {
                    throw new ToolArgumentException($"Invalid arguments for {name}: give `sites` or `preset`, not both");
                }
                if (preset is not null && SitePresets.Find(preset) is null && ContentPresets.Find(preset) is null)
                {
                    throw new ToolArgumentException($"Invalid arguments for {name}: there's no predefined selection \"{preset}\" (list_presets)");
                }
                var items = sites ?? (preset is null ? rows.Select(r => SpNode.SiteFromUrl(r.Url)).DistinctBy(n => n.Key).ToList() : []);
                if (preset is null && items.Count == 0)
                {
                    throw new ToolArgumentException($"Invalid arguments for {name}: give `sites`, `preset` or `rows`");
                }
                // a predefined selection's items are only known when it runs: check the rows' own items then
                var checkOn = preset is null ? items : rows.Select(r => SpNode.SiteFromUrl(r.Url)).ToList();
                var (problems, notes) = PlanCheck.RunProblems(plan, checkOn, values, rows, DateTimeOffset.Now);
                var undeclared = rows.SelectMany(r => r.Values.Keys).Concat(values.Keys).Distinct(StringComparer.OrdinalIgnoreCase)
                    .Where(k => plan.Variables.All(v => !string.Equals(v.Name, k, StringComparison.OrdinalIgnoreCase))).ToList();
                if (undeclared.Count > 0)
                {
                    problems = [.. problems, $"The plan has no variable {string.Join(", ", undeclared)}."];
                }
                if (problems.Count > 0)
                {
                    return new ToolResult("The run wasn't prepared:\n- " + string.Join("\n- ", problems), true);
                }
                var saved = backend.Plans.Save(plan with
                {
                    Prepared = new PreparedRun { Preset = preset, Items = items, Values = values, Rows = rows, By = Client, At = DateTimeOffset.Now },
                });
                return new ToolResult(JsonSerializer.Serialize(new
                {
                    saved.Id,
                    runsOn = Selection(preset, items, rows),
                    notes = notes.Count > 0 ? notes : null,
                    note = $"Ready. The user opens the plan in {AppInfo.Name} (open_app with `plan`), looks it over and runs it; get_plan shows the results afterwards.",
                }, Json));
            }
            case "delete_plan":
            {
                var plan = Plan(a.Required("id"));
                backend.Plans.Delete(plan.Id);
                return new ToolResult(JsonSerializer.Serialize(new { deleted = plan.Id, plan.Name }, Json));
            }
            default:
                return await Task.FromResult<ToolResult?>(null).ConfigureAwait(false);
        }
    }

    private Plan Plan(string id) =>
        backend.Plans.Load(id) ?? throw new SharePointOperationException($"There's no plan \"{id}\" (list_plans lists them).");

    /// <summary>Checks and saves a plan; one with problems isn't saved.</summary>
    private ToolResult Saved(Plan plan, string note)
    {
        if (PlanCheck.Problems(plan) is { Count: > 0 } problems)
        {
            return new ToolResult("The plan wasn't saved:\n- " + string.Join("\n- ", problems), true);
        }
        var saved = backend.Plans.Save(plan);
        return new ToolResult(JsonSerializer.Serialize(new
        {
            saved.Id,
            saved.Name,
            startsOn = PlanCheck.StartKinds(saved),
            variables = saved.Variables.Select(v => v.Name).ToList(),
            steps = saved.Steps.Count,
            note,
        }, Json));
    }

    private static NodeKind? KindOf(Args a, string arg, string tool) =>
        a.String(arg) is { } k
            ? Kinds.Contains(k) ? Enum.Parse<NodeKind>(k, ignoreCase: true) : throw new ToolArgumentException($"Invalid arguments for {tool}: {arg}: one of {string.Join(", ", Kinds)}")
            : null;

    /// <summary>A plan with the test selection the call gives (sites or a preset), else as it was.</summary>
    private static Plan WithTestSelection(Plan plan, Args a, string tool)
    {
        var preset = a.String("preset");
        var sites = SiteNodes(a, "sites", tool);
        if (preset is not null && sites is not null)
        {
            throw new ToolArgumentException($"Invalid arguments for {tool}: give `sites` or `preset`, not both");
        }
        return preset is not null ? plan with { Preset = preset, Items = [], Sites = [], Parents = [] }
            : sites is not null ? plan with { Preset = null, Items = sites, Sites = sites, Parents = [] }
            : plan;
    }

    /// <summary>Site collections' URLs as rows; null when the call gives none.</summary>
    private static List<SpNode>? SiteNodes(Args a, string arg, string tool)
    {
        if (a.Node(arg) is not { } sites)
        {
            return null;
        }
        var nodes = (sites as JsonArray ?? throw new ToolArgumentException($"Invalid arguments for {tool}: {arg}: expected an array of URLs"))
            .Select(n => n?.GetValueKind() == JsonValueKind.String ? n.GetValue<string>().Trim() : throw new ToolArgumentException($"Invalid arguments for {tool}: {arg}: expected URLs"))
            .Select(u => Https(u, tool, arg) ? SpNode.SiteFromUrl(u) : null!)
            .DistinctBy(n => n.Key)
            .ToList();
        return nodes.Count == 0 ? throw new ToolArgumentException($"Invalid arguments for {tool}: {arg}: give at least one") : nodes;
    }

    private static bool Https(string url, string tool, string arg) =>
        Uri.TryCreate(url, UriKind.Absolute, out var uri) && uri.Scheme == Uri.UriSchemeHttps
            ? true
            : throw new ToolArgumentException($"Invalid arguments for {tool}: {arg}: {url} isn't an https URL");

    private static List<PlanVariable>? VariablesOf(Args a, string tool)
    {
        if (a.Node("variables") is not { } node)
        {
            return null;
        }
        return (node as JsonArray ?? throw new ToolArgumentException($"Invalid arguments for {tool}: variables: expected an array")).Select(item =>
        {
            if (item is not JsonObject v || v["name"]?.GetValueKind() != JsonValueKind.String)
            {
                throw new ToolArgumentException($"Invalid arguments for {tool}: variables: each one needs a `name`");
            }
            var type = v["type"]?.GetValue<string>() ?? "text";
            return new PlanVariable
            {
                Name = v["name"]!.GetValue<string>().Trim().Trim('{', '}'),
                Label = v["label"]?.GetValue<string>(),
                Hint = v["hint"]?.GetValue<string>(),
                Kind = VariableTypes.Contains(type) ? Enum.Parse<ParameterKind>(type, ignoreCase: true)
                    : throw new ToolArgumentException($"Invalid arguments for {tool}: variables: type is one of {string.Join(", ", VariableTypes)}"),
            };
        }).ToList();
    }

    /// <summary>The steps a call gives, or null when it gives none.</summary>
    private static List<PlanStep>? StepsOf(Args a)
    {
        if (a.Node("steps") is not { } node)
        {
            return null;
        }
        var steps = new List<PlanStep>();
        foreach (var item in node as JsonArray ?? throw new ToolArgumentException("Invalid arguments: steps: expected an array"))
        {
            if (item is not JsonObject step || step["command"]?.GetValueKind() != JsonValueKind.String)
            {
                throw new ToolArgumentException("Invalid arguments: steps: each one needs a `command`");
            }
            steps.Add(new PlanStep
            {
                Command = step["command"]!.GetValue<string>().Trim(),
                Args = ValuesOf(step["args"] as JsonObject).ToDictionary(x => x.Key, x => (string?)x.Value),
            });
        }
        return steps;
    }

    private static List<PlanRow> RowsOf(Args a, string tool)
    {
        if (a.Node("rows") is not { } node)
        {
            return [];
        }
        return (node as JsonArray ?? throw new ToolArgumentException($"Invalid arguments for {tool}: rows: expected an array")).Select(item =>
        {
            if (item is not JsonObject row || row["url"]?.GetValueKind() != JsonValueKind.String)
            {
                throw new ToolArgumentException($"Invalid arguments for {tool}: rows: each one needs a `url`");
            }
            var url = row["url"]!.GetValue<string>().Trim().TrimEnd('/');
            Https(url, tool, "rows");
            return new PlanRow(url, ValuesOf(row["values"] as JsonObject));
        }).ToList();
    }

    /// <summary>Values by name, as text: strings as they are, numbers invariant, booleans "true"/"false".</summary>
    private static Dictionary<string, string> ValuesOf(JsonObject? given)
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, value) in given ?? [])
        {
            values[key.Trim().Trim('{', '}')] = value?.GetValueKind() switch
            {
                null or JsonValueKind.Null => "",
                JsonValueKind.String => value.GetValue<string>(),
                JsonValueKind.True => "true",
                JsonValueKind.False => "false",
                JsonValueKind.Number => value.AsValue().TryGetValue<long>(out var l) ? l.ToString(CultureInfo.InvariantCulture) : value.ToJsonString(),
                _ => throw new ToolArgumentException($"Invalid arguments: {key} should be a string, number or true/false"),
            };
        }
        return values;
    }

    private static string Selection(string? preset, IReadOnlyList<SpNode> items, IReadOnlyList<PlanRow> rows) =>
        preset is not null
            ? $"predefined selection {preset} ({SitePresets.Find(preset)?.Title ?? ContentPresets.Find(preset)?.Title ?? "unknown"}){(rows.Count > 0 ? $", values for {rows.Count} items" : "")}"
            : $"{string.Join(", ", items.GroupBy(n => n.Kind).Select(g => NodeKinds.Count(g.Key, g.Count())))}{(rows.Count > 0 ? $", values for {rows.Count}" : "")}";

    private static object CommandView(Command c) => new
    {
        c.Id,
        c.Title,
        c.Group,
        c.Description,
        c.PowerShell,
        runsOn = c.Targets,
        finder = c.After == After.Tick ? true : (bool?)null,
        changes = c.ReadOnly ? false : (bool?)null,
        destructive = c.Destructive ? true : (bool?)null,
        needsAdmin = c.NeedsAdmin ? true : (bool?)null,
        parameters = c.Parameters.Select(p => new
        {
            p.Name,
            p.Label,
            type = p.Kind,
            required = p.Required ? true : (bool?)null,
            p.Default,
            p.Min,
            p.Max,
            choices = p.Choices.Count > 0 ? p.Choices.Select(x => new { x.Value, x.Label }) : null,
            p.Hint,
            onlyWhen = p.ShowIf is { } s ? new { parameter = s.Parameter, values = s.Values } : null,
            onlyOn = p.ForKinds,
        }).ToList(),
    };

    private static object RunView(RunRecord r) => new
    {
        r.Started,
        r.Finished,
        complete = r.Finished is not null,
        r.Cancelled,
        done = r.Count(RunOutcome.Done),
        unchanged = r.Count(RunOutcome.Unchanged),
        notRun = r.Count(RunOutcome.NotRun),
        failed = r.Count(RunOutcome.Failed) + r.Count(RunOutcome.Pending),
        failures = r.Targets.Where(t => t.Outcome is RunOutcome.Failed or RunOutcome.Pending).Take(20)
            .Select(t => new { t.Target.Url, t.Message }).ToList(),
    };
}
