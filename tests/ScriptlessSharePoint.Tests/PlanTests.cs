using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.Plans;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Tests;

public sealed class PlanTests : IDisposable
{
    private readonly string folder = Path.Combine(Path.GetTempPath(), "ssp-plans-" + Guid.NewGuid().ToString("N"));

    public void Dispose()
    {
        if (Directory.Exists(folder))
        {
            Directory.Delete(folder, true);
        }
    }

    private const string Tenant = "https://contoso.sharepoint.com";
    private static readonly SpNode Hr = SpNode.SiteFromUrl(Tenant + "/sites/hr");
    private static readonly SpNode Falcon = SpNode.SiteFromUrl(Tenant + "/sites/falcon");
    private static readonly DateTimeOffset Now = new(2026, 11, 10, 12, 0, 0, TimeSpan.Zero);

    private static Plan Plan(params PlanStep[] steps) => new() { Name = "Set up", ConnectionId = "c1", Steps = steps };

    private static PlanStep Step(string command, params (string, string?)[] args) =>
        new() { Command = command, Args = args.ToDictionary(a => a.Item1, a => a.Item2) };

    private static readonly Dictionary<string, string> NoValues = [];

    [Fact]
    public void A_finder_hands_what_it_finds_to_the_next_step()
    {
        // sites → their Site Pages libraries → versioning on them
        var plan = Plan(Step("find.lists", ("which", "sitepages")), Step("list.versioning", ("major", "100"))) with { StartsOn = [NodeKind.SiteCollection] };
        Assert.Empty(PlanCheck.Problems(plan));
    }

    [Fact]
    public void A_plan_starts_on_what_it_says_else_on_what_its_first_step_runs_on()
    {
        Assert.Equal([NodeKind.SiteCollection], PlanCheck.StartKinds(Plan(Step("site.storage", ("max", "100")))));
        Assert.Empty(PlanCheck.StartKinds(Plan(Step("tenant.reshare"))));
        var sites = Plan(Step("list.versioning")) with { StartsOn = [NodeKind.SiteCollection] };
        Assert.Contains(PlanCheck.Problems(sites), p => p.Contains("Step 1 (list.versioning)") && p.Contains("site collections"));
        Assert.True(PlanCheck.StartsOnAny(Plan(Step("web.reindex")), [NodeKind.SiteCollection, NodeKind.File]));
        Assert.False(PlanCheck.StartsOnAny(Plan(Step("web.reindex")), [NodeKind.File]));
        Assert.True(PlanCheck.StartsOnAny(Plan(Step("tenant.reshare")), []));
    }

    [Fact]
    public void Unknown_commands_inputs_and_bad_values_are_problems()
    {
        var problems = PlanCheck.Problems(Plan(
            Step("web.nonsense"),
            Step("web.title", ("title", null), ("colour", "red")),
            Step("site.storage", ("max", "lots"))) with { StartsOn = [NodeKind.SiteCollection] });
        Assert.Contains(problems, p => p.StartsWith("Step 1 (web.nonsense): there's no such command"));
        Assert.Contains(problems, p => p.StartsWith("Step 2 (web.title): it has no input colour"));
        Assert.Contains(problems, p => p.StartsWith("Step 2 (web.title)") && p.Contains("Title"));
        Assert.Contains(problems, p => p.StartsWith("Step 3 (site.storage)"));
        Assert.Contains(PlanCheck.Problems(Plan(Step("web.reindex")) with { Preset = "nope" }), p => p.Contains("no predefined selection"));
    }

    [Fact]
    public void Variables_must_be_declared_and_are_checked_when_their_values_are_known()
    {
        // {quota} in a number input passes until it has a value; {title} is built in
        var plan = Plan(Step("site.storage", ("max", "{quota}")), Step("web.title", ("title", "{title} (archived)"), ("description", "Owner: {owner}")));
        Assert.Contains(PlanCheck.Problems(plan), p => p.Contains("{quota}") && p.Contains("isn't one of the plan's variables"));
        plan = plan with { Variables = [new PlanVariable { Name = "quota", Kind = ParameterKind.Number }, new PlanVariable { Name = "owner", Kind = ParameterKind.Person }] };
        Assert.Empty(PlanCheck.Problems(plan));
        Assert.Contains(PlanCheck.Problems(plan with { Variables = [.. plan.Variables, new PlanVariable { Name = "title" }] }), p => p.Contains("built-in"));
        Assert.Contains(PlanCheck.Problems(plan with { Variables = [.. plan.Variables, new PlanVariable { Name = "2nd" }] }), p => p.Contains("letters, digits"));
    }

    [Fact]
    public void A_run_needs_a_value_for_every_item_and_values_that_fit()
    {
        var plan = Plan(Step("site.storage", ("max", "{quota}"))) with { Variables = [new PlanVariable { Name = "quota", Kind = ParameterKind.Number, Label = "Quota" }] };
        var rows = new List<PlanRow> { new(Hr.Url, new Dictionary<string, string> { ["quota"] = "1024" }) };

        var (missing, _) = PlanCheck.RunProblems(plan, [Hr, Falcon], NoValues, rows, Now);
        Assert.Contains(missing, p => p.StartsWith("Quota: no value for") && p.Contains("falcon"));

        // a value for all fills the gap; the row still wins for its site
        var (none, _) = PlanCheck.RunProblems(plan, [Hr, Falcon], new Dictionary<string, string> { ["quota"] = "2048" }, rows, Now);
        Assert.Empty(none);
        Assert.Equal("1024", PlanValues.For(Hr, new Dictionary<string, string> { ["quota"] = "2048" }, rows, Now)["quota"]);

        var bad = new List<PlanRow> { new(Hr.Url, new Dictionary<string, string> { ["quota"] = "lots" }), new(Falcon.Url, new Dictionary<string, string> { ["quota"] = "1" }) };
        var (wrong, _) = PlanCheck.RunProblems(plan, [Hr, Falcon], NoValues, bad, Now);
        Assert.Single(wrong);
        Assert.Contains("hr", wrong[0]);

        var (_, notes) = PlanCheck.RunProblems(plan, [Hr], NoValues, [.. rows, new(Tenant + "/sites/gone", new Dictionary<string, string> { ["quota"] = "1" })], Now);
        Assert.Contains(notes, n => n.Contains("/sites/gone"));
    }

    [Fact]
    public void Items_found_inside_a_site_take_its_row()
    {
        var library = new SpNode { Kind = NodeKind.Library, Title = "Site Pages", Url = Hr.Url + "/SitePages", WebUrl = Hr.Url, SiteUrl = Hr.Url, ListId = Guid.NewGuid() };
        var rows = new List<PlanRow>
        {
            new(Hr.Url, new Dictionary<string, string> { ["owner"] = "alex@contoso.com" }),
            new(Tenant + "/sites/hrteam", new Dictionary<string, string> { ["owner"] = "not this one" }),
        };
        Assert.Equal("alex@contoso.com", PlanValues.For(library, NoValues, rows, Now)["owner"]);
        Assert.False(PlanValues.Covers(Tenant + "/sites/hr", Tenant + "/sites/hrteam"));
    }

    [Fact]
    public void Built_in_values_and_filling_in()
    {
        var values = PlanValues.BuiltIns(Hr, Now);
        Assert.Equal("hr", values["name"]);
        Assert.Equal(Hr.Url, values["site"]);
        Assert.Equal("2026-11-10", values["today"]);
        Assert.Equal("/sites/hr (archived 2026-11-10) {unknown}", PlanValues.Resolve("{title} (archived {today}) {unknown}", values));

        var step = Step("web.title", ("title", "{title} (archived)"));
        var args = PlanValues.ArgsFor(CommandCatalog.Find("web.title")!, step, values);
        Assert.Equal("/sites/hr (archived)", args["title"]);
    }

    [Theory]
    [InlineData("URL\towner\tquota\nhttps://contoso.sharepoint.com/sites/hr\talex@contoso.com\t1024\n")]
    [InlineData("url,owner,quota\r\n\"https://contoso.sharepoint.com/sites/hr\",alex@contoso.com,1024\r\n")]
    [InlineData("Site;{owner};{quota}\nhttps://contoso.sharepoint.com/sites/hr/;alex@contoso.com;1024")]
    public void Tables_paste_from_excel_or_csv(string text)
    {
        var (rows, columns, problems) = PlanValues.ParseTable(text);
        Assert.Empty(problems);
        Assert.Equal(["owner", "quota"], columns);
        var row = Assert.Single(rows);
        Assert.Equal("https://contoso.sharepoint.com/sites/hr", row.Url);
        Assert.Equal("alex@contoso.com", row.Values["owner"]);
        Assert.Equal("1024", row.Values["QUOTA"]);
    }

    [Fact]
    public void Table_problems_and_round_trip()
    {
        var (_, _, problems) = PlanValues.ParseTable("URL\towner\nhr\talex\nhttps://x.sharepoint.com/sites/a\tsam\nhttps://x.sharepoint.com/sites/a\tkim");
        Assert.Contains(problems, p => p.StartsWith("Row 2"));
        Assert.Contains(problems, p => p.Contains("more than one row"));
        Assert.NotEmpty(PlanValues.ParseTable("just a heading").Problems);

        var rows = new List<PlanRow> { new(Hr.Url, new Dictionary<string, string> { ["owner"] = "alex" }) };
        Assert.Equal(rows[0].Values["owner"], PlanValues.ParseTable(PlanValues.TableText(rows, ["owner"])).Rows[0].Values["owner"]);
    }

    [Fact]
    public async Task The_runner_gives_each_target_its_own_inputs()
    {
        var seen = new List<string>();
        var command = new Command
        {
            Id = "test",
            Title = "Test",
            Group = "Settings",
            Description = "A test.",
            Targets = [NodeKind.SiteCollection],
            Parameters = [new("max", "Max", ParameterKind.Number) { Min = 1 }],
            Run = (c, n) =>
            {
                lock (seen)
                {
                    seen.Add($"{n.Title}={c.Args["max"]}");
                }
                return Task.FromResult(CommandResult.Done("ok"));
            },
        };
        var client = new SharePointClient(new Connections.ConnectionProfile { TenantUrl = Tenant }, new CommandCatalogTests.NoTokens(), false);
        var results = await CommandRunner.RunAsync(command, [Hr, Falcon], new CommandArgs { ["max"] = "{quota}" }, client, new CommandCatalogTests.NoHost(),
            argsFor: n => new CommandArgs { ["max"] = n == Hr ? "10" : "lots" });
        Assert.Equal(["/sites/hr=10"], seen);
        Assert.Contains("number", results.Single(r => r.Target == Falcon).Result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Plans_are_saved_listed_changed_and_deleted()
    {
        var store = new PlanStore(folder);
        var saved = store.Save(Plan(Step("web.reindex")) with
        {
            Items = [Hr],
            Variables = [new PlanVariable { Name = "owner", Kind = ParameterKind.Person }],
            Prepared = new PreparedRun { Items = [Hr], Rows = [new(Hr.Url, new Dictionary<string, string> { ["owner"] = "alex" })] },
        });
        var loaded = store.Load(saved.Id)!;
        Assert.Equal("web.reindex", Assert.Single(loaded.Steps).Command);
        Assert.Equal(Hr.Key, Assert.Single(loaded.Items).Key);
        Assert.Equal(ParameterKind.Person, loaded.Variables[0].Kind);
        Assert.Equal("alex", loaded.Prepared!.Rows[0].Values["owner"]);

        var changed = store.Save(loaded with { Name = "Reindex HR" });
        Assert.Equal(saved.Id, changed.Id);
        Assert.Equal(saved.Created, changed.Created);
        Assert.Equal("Reindex HR", Assert.Single(store.List()).Name);

        store.Delete(saved.Id);
        Assert.Empty(store.List());
    }
}
