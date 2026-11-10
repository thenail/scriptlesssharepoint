using System.Text;

using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.Connections;
using ScriptlessSharePoint.History;
using ScriptlessSharePoint.Plans;
using ScriptlessSharePoint.Selections;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.App;

/// <summary>
/// Jobs: every command or plan the app runs is a job, running in the
/// background (the pane's results watch it, and can be left), kept in
/// History\ in the data folder as it goes (one record per step). Jobs can be
/// paused (the targets running finish, the rest wait), carried on and
/// cancelled. Jobs that were running or paused when the app closed come back
/// paused at start. The Jobs page lists the ones running or paused, then the
/// finished ones (the history): their targets can be ticked again, and a
/// command run again with the same inputs.
/// </summary>
public sealed partial class MainWindow
{
    private const string HistoryTag = "history";

    private readonly RunHistoryStore historyStore = new();

    /// <summary>The jobs not over yet: running or paused, the newest first.</summary>
    private readonly List<Job> jobs = [];

    /// <summary>The page's list: jobs running or paused, then finished ones.</summary>
    private List<JobEntry> entries = [];

    /// <summary>What the page shows.</summary>
    private JobEntry? shownEntry;

    private JobBinding? pageBinding;

    /// <summary>A command to open, with the inputs it had, once the table shows a run's targets.</summary>
    private (Command Command, CommandArgs Args)? pendingForm;

    private readonly HashSet<Job> unsaved = [];
    private DispatcherQueueTimer? saveTimer;

    /// <summary>A line of the page: a job running or paused, or a finished one's records (a plan's: one per step).</summary>
    private sealed record JobEntry(string Id, Job? Live, IReadOnlyList<RunRecord> Records)
    {
        public RunRecord First => Records.OrderBy(r => r.Step ?? 1).First();
        public RunRecord Last => Records.OrderBy(r => r.Step ?? 1).Last();
        public string Title => Live?.Title ?? First.PlanName ?? First.Title;
        public DateTimeOffset Started => Live?.Started ?? First.Started;
    }

    // ---- keeping them ----

    /// <summary>Saves a run, keeping it as it was when the disk says no: the history never stops a run.</summary>
    private RunRecord SaveRun(RunRecord run)
    {
        try
        {
            return historyStore.Save(run);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            Diagnostics.Log.Write("history", $"couldn't save {run.CommandId}: {e.Message}");
            return run;
        }
    }

    /// <summary>Saves a job's step soon: at most once a second, however fast its targets finish.</summary>
    private void SaveSoon(Job job)
    {
        unsaved.Add(job);
        if (saveTimer is null)
        {
            saveTimer = ui.CreateTimer();
            saveTimer.Interval = TimeSpan.FromSeconds(1);
            saveTimer.IsRepeating = false;
            saveTimer.Tick += (_, _) => SaveNow();
        }
        if (!saveTimer.IsRunning)
        {
            saveTimer.Start();
        }
    }

    private void SaveNow()
    {
        foreach (var job in unsaved.ToList())
        {
            if (job.Record is not null)
            {
                job.Record = SaveRun(job.Snapshot());
            }
        }
        unsaved.Clear();
    }

    // ---- starting and running ----

    /// <summary>Starts a command (or, with <paramref name="plan"/>, a plan's steps) as a job on these targets.</summary>
    private Job? StartJob(Command c, List<SpNode> ticked, CommandArgs args, bool ask, Plan? plan = null,
        IReadOnlyDictionary<string, string>? values = null, IReadOnlyList<PlanRow>? rows = null)
    {
        if (active is null)
        {
            return null;
        }
        Session session;
        try
        {
            session = workspace.Current();
        }
        catch (Exception e)
        {
            ShowError(e);
            return null;
        }
        var job = new Job
        {
            Id = $"{DateTime.Now:yyyyMMdd-HHmmss}-{Guid.NewGuid().ToString("N")[..6]}",
            ConnectionId = active.Id,
            ConnectionName = active.Name,
            Plan = plan,
            Ask = ask,
            Values = values ?? new Dictionary<string, string>(),
            Rows = rows ?? [],
            Tenant = Targets(NodeKind.Tenant).FirstOrDefault(),
            Command = c,
            Args = args,
            Ticked = ticked,
        };
        jobs.Insert(0, job);
        _ = RunJobAsync(job, session.Client, prepare: true);
        return job;
    }

    /// <summary>Sets up a job's step: its command and inputs, what it runs on (of what's ticked), and its record. False when it has nothing to run on.</summary>
    private bool PrepareStep(Job job)
    {
        if (job.Plan is { } plan)
        {
            var step = plan.Steps[job.Step - 1];
            if (CommandCatalog.Find(step.Command) is not { } command)
            {
                job.Stopped = $"Step {job.Step}: there's no command {step.Command} any more.";
                return false;
            }
            job.Command = command;
            job.Args = PlanCheck.ArgsOf(command, step);
        }
        var fits = PlanCheck.RunsOn(job.Command, job.Ticked.Select(n => n.Kind).ToHashSet());
        if (fits.Count == 0)
        {
            job.Stopped = $"Step {job.Step} ({job.Command.Title}) has nothing ticked to run on, so it and the steps after it didn't run.";
            return false;
        }
        job.Kind = fits.Contains(NodeKind.Tenant) ? NodeKind.Tenant : fits.MaxBy(k => job.Ticked.Count(n => n.Kind == k));
        var tenant = job.Kind == NodeKind.Tenant;
        var targets = tenant
            ? job.Tenant is { } t ? [t] : []
            : job.Ticked.Where(n => fits.Contains(n.Kind)).DistinctBy(n => n.Key).ToList();
        if (targets.Count == 0)
        {
            job.Stopped = $"Step {job.Step} ({job.Command.Title}) has nothing to run on.";
            return false;
        }
        job.Targets.Clear();
        foreach (var n in targets)
        {
            job.Targets[n.Key] = new JobTarget
            {
                Node = n,
                Site = tenant ? null : siteOf.GetValueOrDefault(n.Key) ?? SpNode.SiteFromUrl(n.SiteUrl ?? n.Url),
                Parent = job.FoundParents.GetValueOrDefault(n.Key) ?? parentOf.GetValueOrDefault(n.Key),
            };
        }
        var inputs = InputsText(job.Command, job.Args);
        job.Record = new RunRecord
        {
            CommandId = job.Command.Id,
            Title = job.Command.Title,
            Plan = job.Plan?.Id,
            PlanName = job.Plan?.Name,
            Step = job.Plan is null ? null : job.Step,
            ConnectionId = job.ConnectionId,
            ConnectionName = job.ConnectionName,
            Kind = job.Kind,
            Started = DateTimeOffset.Now,
            Args = job.Args.ToDictionary().Where(a => a.Value is not null).ToDictionary(a => a.Key, a => a.Value),
            Inputs = inputs.Split('\n', StringSplitOptions.RemoveEmptyEntries),
            JobId = job.Id,
            Ask = job.Ask,
            PlanSnapshot = job.Plan,
            Values = job.Plan is null ? null : job.Values,
            Rows = job.Plan is null ? null : job.Rows,
            Ticked = job.Ticked,
        };
        job.Record = SaveRun(job.Snapshot());
        if (job.Record.Args.Count > 0)
        {
            lastRuns?.Remove(job.Command.Id);
            lastRuns?.Add(job.Command.Id, job.Record);
        }
        job.Notify();
        return true;
    }

    /// <summary>On to a plan's next step: what a finder found, else what was ticked. False at the end (or when a finder found nothing).</summary>
    private static bool Advance(Job job)
    {
        if (job.Plan is null || job.Step >= job.Steps)
        {
            return false;
        }
        if (job.Command.After == After.Tick)
        {
            var found = job.Targets.Values.SelectMany(t => t.Result?.Found ?? []).DistinctBy(f => f.Node.Key).ToList();
            if (found.Count == 0)
            {
                job.Stopped = $"Step {job.Step} ({job.Command.Title}) found nothing, so the steps after it didn't run.";
                return false;
            }
            job.Ticked = found.Select(f => f.Node).ToList();
            foreach (var f in found.Where(f => f.Parent is not null))
            {
                job.FoundParents[f.Node.Key] = f.Parent!;
            }
        }
        job.Step++;
        return true;
    }

    /// <summary>
    /// Runs a job's steps until it's done or cancelled: its waiting targets
    /// (a few at a time, holding back while it's paused, asking first when it
    /// asks), keeping its record as it goes.
    /// </summary>
    private async Task RunJobAsync(Job job, SharePointClient client, bool prepare)
    {
        job.Begin();
        FollowThrottle(job, client.Throttle);
        JobsChanged();
        try
        {
            while (true)
            {
                if (prepare && !PrepareStep(job))
                {
                    break;
                }
                prepare = true;
                var pending = job.Targets.Values.Where(t => t.Phase != TargetPhase.Done).ToList();
                foreach (var t in pending)
                {
                    t.Phase = TargetPhase.Waiting;
                    t.Answer = null;
                }
                job.Notify();

                var progress = new Progress<TargetResult>(r =>
                {
                    if (job.Targets.TryGetValue(r.Target.Key, out var t))
                    {
                        t.Result = r.Result;
                        t.Phase = TargetPhase.Done;
                        t.Answer = null;
                        job.NotifyTarget(t);
                        SaveSoon(job);
                    }
                });
                // the runner calls these from its own threads: the job changes on this one
                void Started(SpNode n) => ui.TryEnqueue(() =>
                {
                    if (job.Targets.TryGetValue(n.Key, out var t) && t.Phase != TargetPhase.Done)
                    {
                        t.Phase = TargetPhase.Running;
                        job.NotifyTarget(t);
                    }
                });
                Task<bool> Approve(SpNode n, CancellationToken ct)
                {
                    var answer = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                    ct.Register(() => answer.TrySetCanceled(ct));
                    ui.TryEnqueue(() =>
                    {
                        if (job.Targets.TryGetValue(n.Key, out var t) && t.Phase != TargetPhase.Done)
                        {
                            t.Phase = TargetPhase.Asking;
                            t.Answer = answer;
                            job.NotifyTarget(t);
                            job.Notify();
                        }
                        else
                        {
                            answer.TrySetResult(false);
                        }
                    });
                    return answer.Task;
                }
                var step = job.Plan?.Steps[job.Step - 1];
                var command = job.Command;
                Func<SpNode, CommandArgs>? argsFor = step is not null && step.Args.Values.Any(v => PlanValues.Used(v).Any())
                    ? n => PlanValues.ArgsFor(command, step, PlanValues.For(n, job.Values, job.Rows, job.Started))
                    : null;

                IReadOnlyList<TargetResult> results;
                try
                {
                    results = await CommandRunner.RunAsync(command, pending.Select(t => t.Node).ToList(), job.Args, client, this, progress, job.Token,
                        Started, job.Ask && !command.ReadOnly ? Approve : null, argsFor, job.WhilePausedAsync);
                }
                catch (Exception e)
                {
                    ShowError(e);
                    results = [];
                    job.Cancel();
                }
                // the run's list is the whole story (progress may still be on its way)
                foreach (var r in results)
                {
                    if (job.Targets.TryGetValue(r.Target.Key, out var t))
                    {
                        t.Result = r.Result;
                        t.Phase = TargetPhase.Done;
                        t.Answer = null;
                    }
                }
                unsaved.Remove(job);
                var cancelled = job.Token.IsCancellationRequested;
                if (cancelled)
                {
                    job.Finish(JobState.Cancelled);
                }
                job.Record = SaveRun(job.Snapshot(DateTimeOffset.Now));
                ApplyJob(job, results);
                if (cancelled)
                {
                    break;
                }
                if (!Advance(job))
                {
                    break;
                }
            }
        }
        finally
        {
            if (!job.IsOver)
            {
                job.Finish(JobState.Done);
            }
            if (job.Record is not null)
            {
                job.Record = SaveRun(job.Snapshot() with { State = job.State });
            }
            jobs.Remove(job);
            JobsChanged();
            if (job.Plan is { } plan)
            {
                ShowToast(job.State == JobState.Cancelled ? $"“{plan.Name}” was cancelled" : job.Stopped is null ? $"Ran “{plan.Name}”" : $"“{plan.Name}” stopped",
                    job.Stopped ?? (job.State == JobState.Cancelled ? $"At step {job.Step}." : $"All {Plural(job.Steps, "step", "steps")} ran: Jobs shows how each went."),
                    ("Jobs", () => Nav.SelectedItem = HistoryItem), success: job.State == JobState.Done && job.Stopped is null);
            }
            else if (watchedJob != job && job.State == JobState.Done)
            {
                ShowToast($"{job.Command.Title}: done", JobSummary(job), ("Jobs", () => Nav.SelectedItem = HistoryItem));
            }
            // a finder that found things, all well: straight on to what to do with them (the table shows
            // them ticked, its message offers Undo, and Jobs keeps how each target went)
            if (watchedJob == job && job.Plan is null && job.Command.After == After.Tick && job.State == JobState.Done &&
                job.ConnectionId == active?.Id && job.Targets.Values.All(t => t.Result is { Ok: true }) &&
                job.Targets.Values.Any(t => t.Result!.Found.Count > 0))
            {
                Unwatch();
                ShowOverview();
            }
            // a report that lists a lot (or over several targets): its table in the table's place, the results stay in the pane
            else if (watchedJob == job && job.Plan is null && job.Command.ReadOnly && job.State == JobState.Done && job.ConnectionId == active?.Id &&
                     job.Targets.Values.Select(t => t.Result?.Table).OfType<ReportTable>().Where(t => t.Rows.Count > 0).ToList() is { Count: > 0 } tables &&
                     (tables.Count > 1 || tables[0].Rows.Count >= 5))
            {
                ShowReport(job.Command.Title, ReportSourcesOf(job));
            }
        }
    }

    /// <summary>What a step did, in the table: renamed and removed rows; for a command (not a plan's step), what a finder found, ticked.</summary>
    private void ApplyJob(Job job, IReadOnlyList<TargetResult> results)
    {
        if (job.ConnectionId != active?.Id || (job.Plan is not null && job.Command.After == After.Tick))
        {
            return;
        }
        var all = job.Targets.Values.Where(t => t.Result is not null).Select(t => new TargetResult(t.Node, t.Result!)).ToList();
        Apply(job.Command, all.Count > 0 ? all : results);
    }

    // ---- SharePoint's throttling: the jobs say when they wait for it ----

    private readonly HashSet<Throttle> followed = [];
    private DispatcherQueueTimer? throttleTimer;

    /// <summary>Follows a job's tenant throttle, so its line counts down while SharePoint has asked everything to wait.</summary>
    private void FollowThrottle(Job job, Throttle throttle)
    {
        job.Throttle = throttle;
        if (followed.Add(throttle))
        {
            // raised on the runner's threads
            throttle.Changed += () => ui.TryEnqueue(OnThrottled);
        }
    }

    private void OnThrottled()
    {
        foreach (var job in jobs.Where(j => j.Live))
        {
            job.Notify();
        }
        if (throttleTimer is null)
        {
            throttleTimer = ui.CreateTimer();
            throttleTimer.Interval = TimeSpan.FromSeconds(1);
            throttleTimer.Tick += (_, _) =>
            {
                var waiting = jobs.Where(j => j.Live && j.Throttle?.PausedUntil is not null).ToList();
                foreach (var job in jobs.Where(j => j.Live))
                {
                    job.Notify();
                }
                if (waiting.Count == 0)
                {
                    throttleTimer!.Stop();
                }
            };
        }
        if (!throttleTimer.IsRunning)
        {
            throttleTimer.Start();
        }
    }

    // ---- pausing, carrying on, cancelling ----

    private void PauseJob(Job job)
    {
        job.Pause();
        SaveSoon(job);
        JobsChanged();
    }

    /// <summary>Carries a job on: a live one's waiting targets go ahead; a restored one starts again from where it was.</summary>
    private async Task ResumeJobAsync(Job job)
    {
        if (job.Live)
        {
            job.Resume();
            SaveSoon(job);
            JobsChanged();
            return;
        }
        // a job from before a restart: its connection, signed in
        if (active?.Id != job.ConnectionId || identity is null)
        {
            if (workspace.Load().Find(job.ConnectionId) is null)
            {
                ShowError(new InvalidOperationException($"\"{job.Title}\" is for a connection that's gone."));
                return;
            }
            await ShowConnectionAsync(job.ConnectionId, force: active?.Id == job.ConnectionId);
            if (active?.Id != job.ConnectionId || identity is null)
            {
                ShowError(new InvalidOperationException($"Sign in to {job.ConnectionName} to carry on with \"{job.Title}\"."));
                return;
            }
            ShowPage(JobsPage);
        }
        job.Tenant ??= Targets(NodeKind.Tenant).FirstOrDefault();
        var prepare = false;
        if (job.StepFinished)
        {
            job.StepFinished = false;
            if (!Advance(job))
            {
                job.Finish(JobState.Done);
                if (job.Record is not null)
                {
                    job.Record = SaveRun(job.Snapshot() with { State = JobState.Done });
                }
                jobs.Remove(job);
                JobsChanged();
                return;
            }
            prepare = true;
        }
        _ = RunJobAsync(job, workspace.Current().Client, prepare);
    }

    private void CancelJob(Job job)
    {
        if (!job.Live)
        {
            foreach (var t in job.Targets.Values.Where(t => t.Phase != TargetPhase.Done))
            {
                t.Result = CommandResult.Failed("Cancelled.");
                t.Phase = TargetPhase.Done;
            }
            job.Cancel();
            if (job.Record is not null)
            {
                job.Record = SaveRun(job.Snapshot(DateTimeOffset.Now));
            }
            jobs.Remove(job);
        }
        else
        {
            job.Cancel();
        }
        JobsChanged();
    }

    /// <summary>
    /// At start: jobs that were running or paused when the app closed, from
    /// their records, paused. A job is its records with one job ID; the last
    /// step's record says where it was.
    /// </summary>
    private void RestoreJobs()
    {
        foreach (var group in historyStore.List().Where(r => r.JobId is not null).GroupBy(r => r.JobId!))
        {
            var last = group.MaxBy(r => r.Step ?? 1)!;
            if (last.State is not (JobState.Running or JobState.Paused) || CommandCatalog.Find(last.CommandId) is not { } command)
            {
                continue;
            }
            var args = new CommandArgs();
            foreach (var (name, value) in last.Args)
            {
                args[name] = value;
            }
            var job = new Job
            {
                Id = group.Key,
                ConnectionId = last.ConnectionId,
                ConnectionName = last.ConnectionName,
                Plan = last.PlanSnapshot,
                Ask = last.Ask,
                Values = last.Values ?? new Dictionary<string, string>(),
                Rows = last.Rows ?? [],
                Tenant = last.Kind == NodeKind.Tenant ? last.Targets.FirstOrDefault()?.Target : null,
                Started = group.Min(r => r.Started),
                Command = command,
                Args = args,
                Kind = last.Kind,
                Ticked = last.Ticked?.ToList() ?? last.Targets.Select(t => t.Target).ToList(),
                Step = last.Step ?? 1,
                Record = last,
                StepFinished = last.Finished is not null,
            };
            foreach (var t in last.Targets)
            {
                job.Targets[t.Target.Key] = new JobTarget
                {
                    Node = t.Target,
                    Site = t.Site,
                    Parent = t.Parent,
                    Phase = t.Outcome == RunOutcome.Pending ? TargetPhase.Waiting : TargetPhase.Done,
                    Result = t.Outcome == RunOutcome.Pending ? null : t.Result(),
                };
            }
            job.Pause();
            job.Live = false;
            job.Record = SaveRun(job.Snapshot());
            jobs.Add(job);
        }
        if (jobs.Count > 0)
        {
            JobsChanged();
            ShowToast($"{Plural(jobs.Count, "job", "jobs")} paused",
                "They were running when the app closed. Carry them on from Jobs.", ("Jobs", () => Nav.SelectedItem = HistoryItem), success: false);
        }
    }

    /// <summary>The jobs changed: the menu's count, and the page when it's showing.</summary>
    private void JobsChanged()
    {
        var live = jobs.Count(j => !j.IsOver);
        HistoryItem.InfoBadge = live > 0 ? new InfoBadge { Value = live } : null;
        if (JobsPage.Visibility == Visibility.Visible)
        {
            LoadHistory(shownEntry?.Id);
        }
    }

    /// <summary>"12 of 40 done · 3 to confirm · paused", or how it ended.</summary>
    private static string JobSummary(Job job)
    {
        var targets = job.Targets.Values.ToList();
        if (!job.IsOver)
        {
            return string.Join(" · ", new[]
            {
                job.Plan is null ? null : job.StepTitle,
                $"{job.DoneCount:N0} of {targets.Count:N0} done",
                job.AskingCount > 0 ? $"{job.AskingCount:N0} to confirm" : null,
                job.Pausing ? "pausing: the ones running finish first" : job.State == JobState.Paused ? job.Live ? "paused" : "paused (the app was closed)" : null,
                job.Live && job.Throttle?.PausedUntil is { } until
                    ? $"SharePoint asked to slow down: carrying on in {Math.Max(1, (int)Math.Ceiling((until - DateTimeOffset.UtcNow).TotalSeconds))} s"
                    : null,
            }.OfType<string>());
        }
        var results = targets.Select(t => t.Result).OfType<CommandResult>().ToList();
        var failed = results.Count(r => !r.Ok);
        var found = results.SelectMany(r => r.Found).Select(f => f.Node).DistinctBy(n => n.Key).ToList();
        return string.Join(" · ", (job.Command.After == After.Tick && job.Plan is null
            ? new[]
            {
                job.State == JobState.Cancelled ? "Cancelled" : null,
                found.Count > 0 ? $"Ticked {Summary(found)}" : "Nothing found",
                failed > 0 ? $"{failed:N0} failed" : null,
            }
            : new[]
            {
                job.State == JobState.Cancelled ? "Cancelled" : null,
                job.Plan is null ? null : job.StepTitle,
                results.Count(r => r is { Ok: true, WasSkipped: false }) is > 0 and var done ? $"{done:N0} done" : null,
                results.Count(r => r is { WasSkipped: true, WasDeclined: false }) is > 0 and var same ? $"{same:N0} unchanged" : null,
                results.Count(r => r.WasDeclined) is > 0 and var declined ? $"{declined:N0} not run" : null,
                failed > 0 ? $"{failed:N0} failed" : null,
            }).OfType<string>());
    }

    // ---- following a job: rows that change as it goes ----

    /// <summary>A list of rows following a job: rebuilt when it moves to another step.</summary>
    private sealed class JobBinding : IDisposable
    {
        private readonly Job job;
        private readonly ItemsControl list;
        private readonly Action changed;
        private readonly Dictionary<string, RunRow> rows = [];
        private int step = -1;
        private int count = -1;

        public JobBinding(Job job, ItemsControl list, Action changed)
        {
            this.job = job;
            this.list = list;
            this.changed = changed;
            job.Changed += OnChanged;
            job.TargetChanged += OnTarget;
            Rebuild();
        }

        public Job Job => job;

        private void Rebuild()
        {
            list.Items.Clear();
            rows.Clear();
            step = job.Step;
            count = job.Targets.Count;
            foreach (var t in job.Targets.Values)
            {
                var target = t;
                var row = new RunRow(t.Node, job.Kind == NodeKind.Tenant, job.Command.Destructive, () => Answered(target));
                rows[t.Node.Key] = row;
                list.Items.Add(row.View);
                Show(t, row);
            }
        }

        private static void Show(JobTarget t, RunRow row)
        {
            switch (t.Phase)
            {
                case TargetPhase.Asking when t.Answer is { } answer:
                    row.Ask(answer);
                    break;
                case TargetPhase.Running:
                    row.Running();
                    break;
                case TargetPhase.Done when t.Result is { } result:
                    row.Show(result);
                    break;
                default:
                    row.Waiting();
                    break;
            }
        }

        /// <summary>A row was answered (Run or Skip): the target waits for its turn, or is skipped.</summary>
        private void Answered(JobTarget t)
        {
            if (t.Phase == TargetPhase.Asking && t.Answer?.Task.IsCompleted == true)
            {
                t.Phase = TargetPhase.Waiting;
                t.Answer = null;
                job.NotifyTarget(t);
                job.Notify();
            }
        }

        private void OnChanged()
        {
            if (job.Step != step || job.Targets.Count != count)
            {
                Rebuild();
            }
            changed();
        }

        private void OnTarget(JobTarget t)
        {
            if (rows.TryGetValue(t.Node.Key, out var row))
            {
                Show(t, row);
            }
            changed();
        }

        /// <summary>Runs every target still waiting to be confirmed.</summary>
        public void RunAll()
        {
            foreach (var t in job.Targets.Values.Where(t => t.Phase == TargetPhase.Asking).ToList())
            {
                t.Answer?.TrySetResult(true);
                Answered(t);
            }
        }

        public void Dispose()
        {
            job.Changed -= OnChanged;
            job.TargetChanged -= OnTarget;
        }
    }

    // ---- the page ----

    private void ShowHistory()
    {
        ShowPage(JobsPage);
        LoadHistory(shownEntry?.Id);
    }

    private void LoadHistory(string? select)
    {
        var live = jobs.Select(j => new JobEntry(j.Id, j, [])).ToList();
        var liveIds = live.Select(e => e.Id).ToHashSet();
        var finished = historyStore.List()
            .GroupBy(r => r.JobId ?? r.File)
            .Where(g => !liveIds.Contains(g.Key))
            .Select(g => new JobEntry(g.Key, null, g.ToList()))
            .OrderByDescending(e => e.Started)
            .ToList();
        entries = [.. live, .. finished];

        HistoryList.SelectionChanged -= OnHistoryPicked;
        HistoryList.Items.Clear();
        var connections = workspace.Load().Connections.ToDictionary(c => c.Id, c => c.Name);
        if (live.Count > 0)
        {
            HistoryList.Items.Add(Heading("Running and paused"));
        }
        foreach (var e in live)
        {
            HistoryList.Items.Add(EntryRow(e, connections));
        }
        if (finished.Count > 0)
        {
            HistoryList.Items.Add(Heading("Finished"));
        }
        foreach (var e in finished)
        {
            HistoryList.Items.Add(EntryRow(e, connections));
        }
        HistoryList.SelectionChanged += OnHistoryPicked;

        HistoryEmpty.Visibility = entries.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        ClearHistoryButton.IsEnabled = finished.Count > 0;
        var picked = entries.FirstOrDefault(e => e.Id == select) ?? entries.FirstOrDefault();
        if (picked is null)
        {
            ShowEntry(null);
            return;
        }
        HistoryList.SelectedItem = HistoryList.Items.OfType<FrameworkElement>().FirstOrDefault(i => i.Tag as string == picked.Id);
        if (!ReferenceEquals(shownEntry?.Live, picked.Live) || shownEntry?.Id != picked.Id || picked.Live is null)
        {
            ShowEntry(picked);
        }
        else
        {
            shownEntry = picked;
            UpdateLiveDetail();
        }

        static FrameworkElement Heading(string text) => new TextBlock
        {
            Text = text.ToUpperInvariant(),
            Style = (Style)Application.Current.Resources["PaneGroupHeader"],
            IsHitTestVisible = false,
        };
    }

    private FrameworkElement EntryRow(JobEntry e, Dictionary<string, string> connections)
    {
        var (glyph, brush) = e.Live is { } job
            ? job.State == JobState.Paused ? ("", "SystemFillColorAttentionBrush") : ("", "AccentFillColorDefaultBrush")
            : Outcome(e.Records);
        var grid = new Grid { ColumnSpacing = 12, Padding = new Thickness(4, 8, 4, 8), Tag = e.Id };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.Children.Add(new FontIcon
        {
            Glyph = glyph,
            FontSize = 16,
            Foreground = (Brush)Application.Current.Resources[brush],
            VerticalAlignment = VerticalAlignment.Top,
            Margin = new Thickness(0, 2, 0, 0),
        });
        var text = new StackPanel { Spacing = 2 };
        var first = e.Live?.Record ?? e.First;
        text.Children.Add(new TextBlock { Text = e.Title, TextTrimming = TextTrimming.CharacterEllipsis });
        text.Children.Add(new TextBlock
        {
            Text = e.Live is { } live ? JobSummary(live) : e.Records.Count > 1 || e.First.PlanName is not null
                ? $"Plan · {Plural(e.Records.Count, "step", "steps")} · {RunSummary(e.Last)}"
                : $"{TargetsText(first?.Kind ?? NodeKind.Tenant, first?.Targets.Count ?? 0)} · {RunSummary(e.Last)}",
            Style = (Style)Application.Current.Resources["CaptionSecondary"],
            TextTrimming = TextTrimming.CharacterEllipsis,
        });
        text.Children.Add(new TextBlock
        {
            Text = $"{When(e.Started)} · {connections.GetValueOrDefault(first?.ConnectionId ?? e.Live?.ConnectionId ?? "") ?? first?.ConnectionName ?? e.Live?.ConnectionName}",
            Style = (Style)Application.Current.Resources["CaptionTextBlockStyle"],
            Foreground = (Brush)Application.Current.Resources["TextFillColorTertiaryBrush"],
            TextTrimming = TextTrimming.CharacterEllipsis,
        });
        Grid.SetColumn(text, 1);
        grid.Children.Add(text);
        return grid;
    }

    /// <summary>A finished job's icon: failed (any target), cut short, all well.</summary>
    private static (string Glyph, string Brush) Outcome(IReadOnlyList<RunRecord> records) =>
        records.Any(r => r.Count(RunOutcome.Failed) > 0) ? ("", "SystemFillColorCriticalBrush")
        : records.Any(r => r.Finished is null || r.Cancelled) ? ("", "SystemFillColorAttentionBrush")
        : ("", "SystemFillColorSuccessBrush");

    /// <summary>"2 done · 1 failed", "Cancelled · 3 done", "Didn't finish".</summary>
    private static string RunSummary(RunRecord r) => string.Join(" · ", new[]
    {
        r.Finished is null ? "Didn't finish" : r.Cancelled ? "Cancelled" : null,
        r.Count(RunOutcome.Done) is > 0 and var done ? $"{done:N0} done" : null,
        r.Count(RunOutcome.Unchanged) is > 0 and var same ? $"{same:N0} unchanged" : null,
        r.Count(RunOutcome.NotRun) is > 0 and var skipped ? $"{skipped:N0} not run" : null,
        r.Count(RunOutcome.Failed) is > 0 and var failed ? $"{failed:N0} failed" : null,
    }.OfType<string>().DefaultIfEmpty("Nothing run"));

    /// <summary>"Today 12:05", "Yesterday 09:14", "3 Nov 2026 16:40".</summary>
    private static string When(DateTimeOffset at)
    {
        var local = at.LocalDateTime;
        var day = (DateTime.Today - local.Date).Days switch
        {
            0 => "Today",
            1 => "Yesterday",
            _ => local.ToString("d MMM yyyy", System.Globalization.CultureInfo.CurrentCulture),
        };
        return $"{day} {local:HH:mm}";
    }

    private static string Took(TimeSpan t) =>
        t.TotalSeconds < 1 ? "under a second" : t.TotalMinutes < 1 ? $"{t.TotalSeconds:0} s" : t.TotalHours < 1 ? $"{(int)t.TotalMinutes} min {t.Seconds} s" : $"{(int)t.TotalHours} h {t.Minutes} min";

    private void OnHistoryPicked(object sender, SelectionChangedEventArgs e)
    {
        if (HistoryList.SelectedItem is FrameworkElement { Tag: string id } && entries.FirstOrDefault(x => x.Id == id) is { } entry)
        {
            ShowEntry(entry);
        }
    }

    /// <summary>A job on the page: a live one follows the job; a finished one shows its records.</summary>
    private void ShowEntry(JobEntry? e)
    {
        pageBinding?.Dispose();
        pageBinding = null;
        shownEntry = e;
        RunDetail.Visibility = e is null ? Visibility.Collapsed : Visibility.Visible;
        RunTargets.Items.Clear();
        if (e is null)
        {
            return;
        }
        if (e.Live is { } job)
        {
            FinishedJobActions.Visibility = Visibility.Collapsed;
            LiveJobActions.Visibility = Visibility.Visible;
            DeleteRunButton.IsEnabled = false;
            pageBinding = new JobBinding(job, RunTargets, UpdateLiveDetail);
            UpdateLiveDetail();
            return;
        }

        FinishedJobActions.Visibility = Visibility.Visible;
        LiveJobActions.Visibility = Visibility.Collapsed;
        JobBar.Visibility = Visibility.Collapsed;
        DeleteRunButton.IsEnabled = true;
        var first = e.First;
        var last = e.Last;
        var connection = workspace.Load().Find(first.ConnectionId);
        var plan = first.PlanName is not null;
        RunTitle.Text = e.Title;
        RunWhen.Text = $"{When(first.Started)}{(last.Finished is { } end ? $" · took {Took(end - first.Started)}" : "")} · {connection?.Name ?? $"{first.ConnectionName} (connection removed)"}";
        RunInputs.Text = plan ? "" : string.Join("\n", first.Inputs);
        RunInputs.Visibility = !plan && first.Inputs.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        RunSummaryText.Text = plan
            ? $"Plan · {Plural(e.Records.Count, "step", "steps")} ran · last: {RunSummary(last)}"
            : $"On {TargetsText(first.Kind, first.Targets.Count)} · {RunSummary(first)}";

        var tenant = first.Kind == NodeKind.Tenant;
        var hasCommand = CommandCatalog.Find(first.CommandId) is not null;
        RunAgainButton.Content = plan ? "Open the plan" : "Run again…";
        RunAgainButton.IsEnabled = connection is not null && (plan ? planStore.Load(first.Plan ?? "") is not null : hasCommand);
        ToolTipService.SetToolTip(RunAgainButton, plan
            ? "Opens the plan, to run it again"
            : !hasCommand ? "This command isn't in the app any more."
            : tenant ? "Opens the command with the same inputs." : "Ticks the same targets and opens the command with the same inputs.");
        TickRunButton.Visibility = tenant ? Visibility.Collapsed : Visibility.Visible;
        TickRunButton.IsEnabled = connection is not null;
        var failed = e.Records.Sum(r => r.Targets.Count(t => t.Outcome is RunOutcome.Failed or RunOutcome.Pending));
        TickRunFailedButton.Visibility = !tenant && failed > 0 && failed < e.Records.Sum(r => r.Targets.Count) ? Visibility.Visible : Visibility.Collapsed;
        TickRunFailedButton.IsEnabled = connection is not null;
        ReportRunButton.Visibility = ReportRecord(e.Records) is not null ? Visibility.Visible : Visibility.Collapsed;
        ReportRunButton.IsEnabled = connection is not null;

        foreach (var r in e.Records.OrderBy(r => r.Step ?? 1))
        {
            if (plan)
            {
                RunTargets.Items.Add(new TextBlock
                {
                    Text = $"Step {r.Step}: {r.Title} · {RunSummary(r)}",
                    Style = (Style)Application.Current.Resources["BodyStrongTextBlockStyle"],
                    Margin = new Thickness(12, 10, 12, 2),
                    TextWrapping = TextWrapping.Wrap,
                });
            }
            foreach (var t in r.Targets)
            {
                var row = new RunRow(t.Target, r.Kind == NodeKind.Tenant, false, () => { });
                row.Show(t.Outcome == RunOutcome.Pending ? CommandResult.Failed(t.Message) : t.Result());
                RunTargets.Items.Add(row.View);
            }
        }
    }

    /// <summary>A live job's heading, progress and buttons, as it goes.</summary>
    private void UpdateLiveDetail()
    {
        if (shownEntry?.Live is not { } job)
        {
            return;
        }
        RunTitle.Text = job.Title;
        RunWhen.Text = $"Started {When(job.Started)} · {job.ConnectionName}";
        var inputs = InputsText(job.Command, job.Args);
        RunInputs.Text = inputs;
        RunInputs.Visibility = inputs.Length > 0 ? Visibility.Visible : Visibility.Collapsed;
        RunSummaryText.Text = JobSummary(job);
        JobBar.Visibility = job.IsOver ? Visibility.Collapsed : Visibility.Visible;
        JobBar.Maximum = Math.Max(1, job.Targets.Count);
        JobBar.Value = job.DoneCount;
        JobBar.ShowPaused = job.State == JobState.Paused;
        JobPauseButton.Content = job.State == JobState.Paused ? "Carry on" : "Pause";
        JobPauseButton.IsEnabled = !job.IsOver && !job.Pausing;
        JobCancelButton.IsEnabled = !job.IsOver;
        JobRunAllButton.Visibility = job.AskingCount > 0 ? Visibility.Visible : Visibility.Collapsed;
        if (job.IsOver)
        {
            // over: it's in the finished list now
            ui.TryEnqueue(() => LoadHistory(job.Id));
        }
    }

    private void OnJobPause(object sender, RoutedEventArgs e)
    {
        if (shownEntry?.Live is not { } job)
        {
            return;
        }
        if (job.State == JobState.Paused)
        {
            _ = ResumeJobAsync(job);
        }
        else
        {
            PauseJob(job);
        }
        UpdateLiveDetail();
    }

    private async void OnJobCancel(object sender, RoutedEventArgs e)
    {
        if (shownEntry?.Live is not { } job)
        {
            return;
        }
        if (await ConfirmAsync($"Cancel “{job.Title}”?", $"What hasn't run yet won't. {job.DoneCount:N0} of {job.Targets.Count:N0} have run{(job.Plan is null ? "" : $" in step {job.Step} of {job.Steps}")}.", "Cancel the job", danger: false))
        {
            CancelJob(job);
        }
    }

    private void OnJobRunAll(object sender, RoutedEventArgs e) => pageBinding?.RunAll();

    // ---- from a finished job ----

    private void OnRunAgain(object sender, RoutedEventArgs e)
    {
        if (shownEntry is not { Live: null } entry)
        {
            return;
        }
        if (entry.First.Plan is { } planId)
        {
            _ = OpenPlanAsync(planId);
            return;
        }
        if (CommandCatalog.Find(entry.First.CommandId) is { } c)
        {
            _ = OpenRunAsync(entry.First, failedOnly: false, c);
        }
    }

    private void OnTickRun(object sender, RoutedEventArgs e)
    {
        if (shownEntry is { Live: null } entry)
        {
            _ = OpenRunAsync(entry.First, failedOnly: false, null);
        }
    }

    private void OnTickRunFailed(object sender, RoutedEventArgs e)
    {
        if (shownEntry is { Live: null } entry)
        {
            var failed = entry.Records.OrderBy(r => r.Step ?? 1).FirstOrDefault(r => r.Targets.Any(t => t.Outcome is RunOutcome.Failed or RunOutcome.Pending)) ?? entry.First;
            _ = OpenRunAsync(failed, failedOnly: true, null);
        }
    }

    /// <summary>
    /// Shows a run's targets ticked in its connection's table (as a saved
    /// selection shows: each under the row it was ticked in, nothing loaded),
    /// then, with <paramref name="again"/>, opens that command with the same inputs.
    /// </summary>
    private async Task OpenRunAsync(RunRecord r, bool failedOnly, Command? again)
    {
        if (workspace.Load().Find(r.ConnectionId) is null)
        {
            return;
        }
        var args = new CommandArgs();
        foreach (var (name, value) in r.Args)
        {
            args[name] = value;
        }
        if (r.Kind == NodeKind.Tenant)
        {
            // nothing to tick: the connection, and the command for the tenant
            SelectNavQuietly(r.ConnectionId);
            await ShowConnectionAsync(r.ConnectionId);
            if (again is not null)
            {
                ShowPane(true);
                pickedKind = NodeKind.Tenant;
                UpdatePane();
                OpenForm(again, args);
            }
            return;
        }

        var targets = r.Targets.Where(t => !failedOnly || t.Outcome is RunOutcome.Failed or RunOutcome.Pending).ToList();
        var s = new SavedSelection
        {
            Name = $"{(failedOnly ? "Failed in " : "")}{r.Title}, {When(r.Started)}",
            ConnectionId = r.ConnectionId,
            Saved = r.Started,
            Items = targets.Select(t => t.Target).ToList(),
            Sites = targets.Select(t => t.Site).OfType<SpNode>().DistinctBy(n => n.Key).ToList(),
            Parents = targets.Select(t => t.Parent).OfType<SpNode>().DistinctBy(n => n.Key).ToList(),
        };
        pendingForm = again is null ? null : (again, args);
        if (active?.Id != r.ConnectionId || identity is null)
        {
            pendingSaved = s;
            SelectNavQuietly(r.ConnectionId);
            await ShowConnectionAsync(r.ConnectionId, force: active?.Id == r.ConnectionId);
            return;
        }
        SelectNavQuietly(r.ConnectionId);
        ShowSaved(s);
    }

    /// <summary>After a run's targets show: the command it ran, with its inputs, when asked for.</summary>
    private void OpenPendingForm()
    {
        if (pendingForm is { } f)
        {
            pendingForm = null;
            if (f.Command.AppliesTo(pickedKind) && Targets(pickedKind).Count > 0)
            {
                OpenForm(f.Command, f.Args);
            }
        }
        pendingForm = null;
    }

    private void OnCopyRun(object sender, RoutedEventArgs e)
    {
        if (shownEntry is not { } entry)
        {
            return;
        }
        var text = new StringBuilder();
        if (entry.Live is { } job)
        {
            text.AppendLine($"{job.Title} · {job.Started.LocalDateTime:g} · {job.ConnectionName} · {JobSummary(job)}");
            foreach (var t in job.Targets.Values)
            {
                text.AppendLine($"{(t.Result is null ? "…" : t.Result.Ok ? t.Result.WasSkipped ? "–" : "✓" : "✗")} {t.Node.Url}: {t.Result?.Message ?? t.Phase.ToString()}");
            }
            CopyText(text.ToString());
            return;
        }
        foreach (var r in entry.Records.OrderBy(r => r.Step ?? 1))
        {
            text.AppendLine($"{(r.Step is { } s ? $"Step {s}: " : "")}{r.Title} · {r.Started.LocalDateTime:g} · {r.ConnectionName}");
            foreach (var line in r.Inputs)
            {
                text.AppendLine($"  {line}");
            }
            foreach (var t in r.Targets)
            {
                var mark = t.Outcome switch { RunOutcome.Done => "✓", RunOutcome.Failed or RunOutcome.Pending => "✗", _ => "–" };
                text.AppendLine($"{mark} {t.Target.Url}: {t.Message}");
                if (t.Detail is { Length: > 0 } d)
                {
                    text.AppendLine(string.Join(Environment.NewLine, d.Split('\n').Select(l => "    " + l.TrimEnd('\r'))));
                }
            }
        }
        CopyText(text.ToString());
    }

    private void OnDeleteRun(object sender, RoutedEventArgs e)
    {
        if (shownEntry is not { Live: null } entry)
        {
            return;
        }
        var i = entries.FindIndex(x => x.Id == entry.Id);
        try
        {
            foreach (var r in entry.Records)
            {
                historyStore.Delete(r.File);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return;
        }
        shownEntry = null;
        LoadHistory(i >= 0 && i + 1 < entries.Count ? entries[i + 1].Id : null);
    }

    private async void OnClearHistory(object sender, RoutedEventArgs e)
    {
        if (!await ConfirmAsync("Clear the finished jobs?", "Every finished job goes from the list. Nothing changes in SharePoint, and jobs running or paused stay.", "Clear", danger: false))
        {
            return;
        }
        var keep = jobs.Select(j => j.Id).ToHashSet();
        try
        {
            foreach (var r in historyStore.List().Where(r => r.JobId is null || !keep.Contains(r.JobId)))
            {
                historyStore.Delete(r.File);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
        }
        shownEntry = null;
        LoadHistory(null);
    }
}
