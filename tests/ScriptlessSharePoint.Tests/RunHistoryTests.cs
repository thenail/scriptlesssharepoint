using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.History;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Tests;

public sealed class RunHistoryTests : IDisposable
{
    private readonly string folder = Path.Combine(Path.GetTempPath(), "ssp-history-" + Guid.NewGuid().ToString("N"));

    public void Dispose()
    {
        if (Directory.Exists(folder))
        {
            Directory.Delete(folder, true);
        }
    }

    private static readonly SpNode Docs = new()
    {
        Kind = NodeKind.Library,
        Title = "Documents",
        Url = "https://contoso.sharepoint.com/sites/hr/Shared Documents",
        WebUrl = "https://contoso.sharepoint.com/sites/hr",
        SiteUrl = "https://contoso.sharepoint.com/sites/hr",
        ListId = Guid.NewGuid(),
    };

    private static RunRecord Run(DateTimeOffset started, RunOutcome outcome = RunOutcome.Pending) => new()
    {
        CommandId = "list.reindex",
        Title = "Reindex",
        ConnectionId = "c1",
        ConnectionName = "Contoso",
        Kind = NodeKind.Library,
        Started = started,
        Args = new Dictionary<string, string?> { ["show"] = "true" },
        Inputs = ["Show in search: on"],
        Targets = [new RunTarget { Target = Docs, Site = SpNode.SiteFromUrl(Docs.SiteUrl!), Outcome = outcome, Message = "Waiting…" }],
    };

    [Fact]
    public void Saves_a_run_as_it_goes_and_lists_the_newest_first()
    {
        var store = new RunHistoryStore(folder);
        var now = DateTimeOffset.Now;
        var older = store.Save(Run(now.AddMinutes(-5)));
        var run = store.Save(Run(now));
        Assert.NotEqual(older.File, run.File);

        // the same file again when it ends
        var ended = store.Save(run with
        {
            Finished = now.AddSeconds(3),
            Targets = [run.Targets[0] with { Outcome = RunOutcome.Done, Message = "Search will crawl it again." }],
        });
        Assert.Equal(run.File, ended.File);

        var list = store.List();
        Assert.Equal([run.File, older.File], list.Select(r => r.File));
        var loaded = list[0];
        Assert.Equal(1, loaded.Count(RunOutcome.Done));
        Assert.Equal(Docs.Key, loaded.Targets[0].Target.Key);
        Assert.Equal("true", loaded.Args["show"]);
        Assert.NotNull(loaded.Finished);
        Assert.Null(list[1].Finished);
    }

    [Fact]
    public void Keeps_the_newest_runs_only()
    {
        var store = new RunHistoryStore(folder);
        var start = DateTimeOffset.Now.AddDays(-1);
        for (var i = 0; i < RunHistoryStore.Keep + 3; i++)
        {
            store.Save(Run(start.AddSeconds(i)));
        }
        var list = store.List();
        Assert.Equal(RunHistoryStore.Keep, list.Count);
        Assert.Equal(start.AddSeconds(RunHistoryStore.Keep + 2), list[0].Started);
    }

    [Fact]
    public void Cuts_long_details_and_deletes()
    {
        var store = new RunHistoryStore(folder);
        var run = Run(DateTimeOffset.Now);
        var saved = store.Save(run with { Targets = [run.Targets[0] with { Detail = new string('x', RunHistoryStore.DetailLimit * 2) }] });
        Assert.True(store.Load(saved.File)!.Targets[0].Detail!.Length < RunHistoryStore.DetailLimit + 10);
        store.Delete(saved.File);
        Assert.Empty(store.List());
    }

    [Theory]
    [InlineData(true, false, false, RunOutcome.Done)]
    [InlineData(true, true, false, RunOutcome.Unchanged)]
    [InlineData(true, true, true, RunOutcome.NotRun)]
    [InlineData(false, false, false, RunOutcome.Failed)]
    public void Outcomes_follow_the_results(bool ok, bool skipped, bool declined, RunOutcome outcome) =>
        Assert.Equal(outcome, RunTarget.OutcomeOf(new CommandResult(ok, "") { WasSkipped = skipped, WasDeclined = declined }));
}
