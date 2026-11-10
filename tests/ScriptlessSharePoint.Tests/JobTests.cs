using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.History;
using ScriptlessSharePoint.Plans;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Tests;

public sealed class JobTests : IDisposable
{
    private readonly string folder = Path.Combine(Path.GetTempPath(), "ssp-jobs-" + Guid.NewGuid().ToString("N"));

    public void Dispose()
    {
        if (Directory.Exists(folder))
        {
            Directory.Delete(folder, true);
        }
    }

    private static readonly SpNode Hr = SpNode.SiteFromUrl("https://contoso.sharepoint.com/sites/hr");
    private static readonly SpNode Falcon = SpNode.SiteFromUrl("https://contoso.sharepoint.com/sites/falcon");

    [Fact]
    public void A_job_step_keeps_what_it_takes_to_carry_on()
    {
        var library = new SpNode { Kind = NodeKind.Library, Title = "Site Pages", Url = Hr.Url + "/SitePages", WebUrl = Hr.Url, SiteUrl = Hr.Url, ListId = Guid.NewGuid() };
        var store = new RunHistoryStore(folder);
        var saved = store.Save(new RunRecord
        {
            CommandId = "find.lists",
            Title = "Tick a library or list in each",
            ConnectionId = "c1",
            Started = DateTimeOffset.Now,
            JobId = "job-1",
            State = JobState.Paused,
            Ask = true,
            Plan = "plan-1",
            Step = 1,
            PlanSnapshot = new Plan { Name = "Versions", ConnectionId = "c1", Steps = [new PlanStep { Command = "find.lists" }, new PlanStep { Command = "list.versioning" }] },
            Values = new Dictionary<string, string> { ["owner"] = "alex" },
            Rows = [new PlanRow(Hr.Url, new Dictionary<string, string> { ["owner"] = "sam" })],
            Ticked = [Hr, Falcon],
            Targets =
            [
                new RunTarget { Target = Hr, Outcome = RunOutcome.Done, Message = "Site Pages", Found = [new FoundNode(library, null)] },
                new RunTarget { Target = Falcon, Outcome = RunOutcome.Pending, Message = "Paused." },
            ],
        });
        var loaded = store.Load(saved.File)!;
        Assert.Equal(("job-1", JobState.Paused, true), (loaded.JobId, loaded.State, loaded.Ask));
        Assert.Equal(2, loaded.PlanSnapshot!.Steps.Count);
        Assert.Equal("sam", loaded.Rows![0].Values["owner"]);
        Assert.Equal([Hr.Key, Falcon.Key], loaded.Ticked!.Select(n => n.Key));

        // a finder's finds come back with its result, for the next step
        var done = loaded.Targets[0].Result();
        Assert.True(done.Ok);
        Assert.Equal(library.Key, Assert.Single(done.Found).Node.Key);
        Assert.Equal(RunOutcome.Pending, loaded.Targets[1].Outcome);
    }

    [Theory]
    [InlineData(RunOutcome.Done, true, false, false)]
    [InlineData(RunOutcome.Unchanged, true, true, false)]
    [InlineData(RunOutcome.NotRun, true, true, true)]
    [InlineData(RunOutcome.Failed, false, false, false)]
    public void Outcomes_give_their_results_back(RunOutcome outcome, bool ok, bool skipped, bool declined)
    {
        var r = new RunTarget { Target = Hr, Outcome = outcome, Message = "m" }.Result();
        Assert.Equal((ok, skipped, declined), (r.Ok, r.WasSkipped, r.WasDeclined));
        Assert.Equal(outcome, RunTarget.OutcomeOf(r));
    }

    private static Command Counting(List<string> ran) => new()
    {
        Id = "test",
        Title = "Test",
        Group = "Settings",
        Description = "A test.",
        Targets = [NodeKind.SiteCollection],
        Run = (_, n) =>
        {
            lock (ran)
            {
                ran.Add(n.Title);
            }
            return Task.FromResult(CommandResult.Done("ok"));
        },
    };

    private static SharePointClient Client() =>
        new(new Connections.ConnectionProfile { TenantUrl = "https://contoso.sharepoint.com" }, new CommandCatalogTests.NoTokens(), false);

    [Fact]
    public async Task A_paused_job_holds_its_targets_until_it_carries_on()
    {
        var ran = new List<string>();
        var resume = new TaskCompletionSource();
        var run = CommandRunner.RunAsync(Counting(ran), [Hr, Falcon], new CommandArgs(), Client(), new CommandCatalogTests.NoHost(),
            beforeStart: ct => resume.Task.WaitAsync(ct));
        await Task.Delay(100);
        Assert.Empty(ran);
        resume.SetResult();
        var results = await run;
        Assert.Equal(2, ran.Count);
        Assert.All(results, r => Assert.True(r.Result.Ok));
    }

    [Fact]
    public async Task Cancelling_a_paused_job_cancels_what_waits()
    {
        var ran = new List<string>();
        using var cts = new CancellationTokenSource();
        var run = CommandRunner.RunAsync(Counting(ran), [Hr, Falcon], new CommandArgs(), Client(), new CommandCatalogTests.NoHost(), ct: cts.Token,
            beforeStart: ct => Task.Delay(Timeout.Infinite, ct));
        await Task.Delay(50);
        await cts.CancelAsync();
        var results = await run;
        Assert.Empty(ran);
        Assert.All(results, r => Assert.Equal("Cancelled.", r.Result.Message));
    }
}
