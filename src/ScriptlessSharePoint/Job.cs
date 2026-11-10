using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.History;
using ScriptlessSharePoint.Plans;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.App;

/// <summary>Where one target of a job's step is.</summary>
internal enum TargetPhase
{
    Waiting,
    Asking,
    Running,
    Done,
}

internal sealed class JobTarget
{
    public required SpNode Node { get; init; }
    public SpNode? Site { get; init; }
    public SpNode? Parent { get; init; }
    public TargetPhase Phase { get; set; }
    public CommandResult? Result { get; set; }

    /// <summary>While it asks Run or Skip: the answer to give.</summary>
    public TaskCompletionSource<bool>? Answer { get; set; }
}

/// <summary>
/// A command, or a plan's steps, running in the background: what's ticked,
/// each target's phase and result, and whether it runs, is paused, or is over.
/// It's kept as run records (one per step) as it goes, so it shows in the
/// history and, if the app closes, comes back paused. Views (the pane's
/// results, the Jobs page) follow it through <see cref="Changed"/> and <see cref="TargetChanged"/>.
/// </summary>
internal sealed class Job
{
    public required string Id { get; init; }
    public required string ConnectionId { get; init; }
    public required string ConnectionName { get; init; }

    /// <summary>The plan it runs, or null for a single command.</summary>
    public Plan? Plan { get; init; }

    public bool Ask { get; init; }
    public IReadOnlyDictionary<string, string> Values { get; init; } = new Dictionary<string, string>();
    public IReadOnlyList<PlanRow> Rows { get; init; } = [];

    /// <summary>The tenant's row, for steps that run on the tenant.</summary>
    public SpNode? Tenant { get; set; }

    /// <summary>Its tenant's throttle: while SharePoint has asked everything to wait, the job says so.</summary>
    public Throttle? Throttle { get; set; }

    /// <summary>The rows found items were found under (a plan's later steps keep them with their items).</summary>
    public Dictionary<string, SpNode> FoundParents { get; } = [];

    public DateTimeOffset Started { get; init; } = DateTimeOffset.Now;

    /// <summary>The step running (from 1).</summary>
    public int Step { get; set; } = 1;

    public int Steps => Plan?.Steps.Count ?? 1;

    public required Command Command { get; set; }
    public required CommandArgs Args { get; set; }
    public NodeKind Kind { get; set; }

    /// <summary>What was ticked as the step started.</summary>
    public List<SpNode> Ticked { get; set; } = [];

    /// <summary>The step's targets, in order, by key.</summary>
    public OrderedDictionary<string, JobTarget> Targets { get; } = new();

    /// <summary>The step's record, kept in the history.</summary>
    public RunRecord? Record { get; set; }

    public JobState State { get; private set; } = JobState.Running;

    /// <summary>Whether a runner is working on it (a job restored after a restart isn't until it's resumed).</summary>
    public bool Live { get; set; }

    /// <summary>Restored with its step over: resuming starts the next one.</summary>
    public bool StepFinished { get; set; }

    /// <summary>Why a plan stopped before its last step.</summary>
    public string? Stopped { get; set; }

    public string Title => Plan is { } p ? p.Name : Command.Title;

    /// <summary>"Step 2 of 3: Versioning", or the command's title.</summary>
    public string StepTitle => Plan is null ? Command.Title : $"Step {Step} of {Steps}: {Command.Title}";

    public int DoneCount => Targets.Values.Count(t => t.Phase == TargetPhase.Done);
    public int AskingCount => Targets.Values.Count(t => t.Phase == TargetPhase.Asking);
    public bool IsOver => State is JobState.Done or JobState.Cancelled;

    /// <summary>Pausing: asked to pause, with targets still finishing.</summary>
    public bool Pausing => State == JobState.Paused && Targets.Values.Any(t => t.Phase == TargetPhase.Running);

    public event Action? Changed;
    public event Action<JobTarget>? TargetChanged;

    private CancellationTokenSource cts = new();
    private TaskCompletionSource resume = Open();

    public CancellationToken Token => cts.Token;

    public void Notify() => Changed?.Invoke();

    public void NotifyTarget(JobTarget t) => TargetChanged?.Invoke(t);

    /// <summary>For a runner starting on it (again).</summary>
    public void Begin()
    {
        if (cts.IsCancellationRequested)
        {
            cts.Dispose();
            cts = new CancellationTokenSource();
        }
        State = JobState.Running;
        resume.TrySetResult();
        Live = true;
    }

    /// <summary>Stops starting targets; those running finish.</summary>
    public void Pause()
    {
        if (State != JobState.Running)
        {
            return;
        }
        State = JobState.Paused;
        resume = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        Notify();
    }

    /// <summary>Carries on: a live job's waiting targets go ahead (a restored one is started by its owner).</summary>
    public void Resume()
    {
        if (State != JobState.Paused)
        {
            return;
        }
        State = JobState.Running;
        resume.TrySetResult();
        Notify();
    }

    public void Cancel()
    {
        cts.Cancel();
        resume.TrySetResult();
        foreach (var t in Targets.Values)
        {
            t.Answer?.TrySetCanceled();
        }
        if (!Live)
        {
            Finish(JobState.Cancelled);
        }
        Notify();
    }

    public void Finish(JobState state)
    {
        State = state;
        Live = false;
        Notify();
    }

    /// <summary>Waits while it's paused.</summary>
    public Task WhilePausedAsync(CancellationToken ct) => State == JobState.Paused ? resume.Task.WaitAsync(ct) : Task.CompletedTask;

    /// <summary>The step as a record for the history: its targets' outcomes so far, and the job's state.</summary>
    public RunRecord Snapshot(DateTimeOffset? finished = null) => Record! with
    {
        State = State,
        Finished = finished ?? Record!.Finished,
        Cancelled = State == JobState.Cancelled,
        Targets = Targets.Values.Select(t => new RunTarget
        {
            Target = t.Node,
            Site = t.Site,
            Parent = t.Parent,
            Outcome = t.Result is { } r ? RunTarget.OutcomeOf(r) : RunOutcome.Pending,
            Message = t.Result?.Message ?? (State == JobState.Paused ? "Paused." : "Didn't finish: the app closed while it ran."),
            Detail = t.Result?.Detail,
            Found = t.Result?.Found is { Count: > 0 } found ? found : null,
            Table = t.Result?.Table,
        }).ToList(),
    };

    private static TaskCompletionSource Open()
    {
        var done = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        done.SetResult();
        return done;
    }
}
