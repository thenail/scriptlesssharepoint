using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Commands;

/// <summary>How one target of a run went.</summary>
public sealed record TargetResult(SpNode Target, CommandResult Result);

/// <summary>Runs a command on many targets: a few at a time, each on its own, reporting as each finishes.</summary>
public static class CommandRunner
{
    /// <summary>How many targets run at once: enough to be quick, few enough not to be throttled.</summary>
    public const int Parallel = 4;

    /// <summary>
    /// Runs <paramref name="command"/> on each target it applies to. A target's
    /// failure is its result, not the run's; cancelling stops targets not yet started.
    /// <paramref name="started"/> hears of each target as it starts.
    /// With <paramref name="approve"/>, each target waits for it (in any order,
    /// before taking its turn) and is <see cref="CommandResult.Declined"/> on no.
    /// </summary>
    public static async Task<IReadOnlyList<TargetResult>> RunAsync(
        Command command,
        IReadOnlyList<SpNode> targets,
        CommandArgs args,
        SharePointClient client,
        ICommandHost host,
        IProgress<TargetResult>? progress = null,
        CancellationToken ct = default,
        Action<SpNode>? started = null,
        Func<SpNode, CancellationToken, Task<bool>>? approve = null)
    {
        if (command.Check(args, targets.Where(t => command.AppliesTo(t.Kind)).Select(t => t.Kind)) is { Count: > 0 } problems)
        {
            throw new ArgumentException(string.Join(" ", problems));
        }
        var applicable = targets.Where(t => command.AppliesTo(t.Kind)).DistinctBy(t => t.Key).ToList();
        var results = new TargetResult[applicable.Count];
        using var gate = new SemaphoreSlim(command.OneAtATime ? 1 : Parallel);
        var context = new CommandContext(client, args, host, ct);

        var tasks = applicable.Select(async (target, i) =>
        {
            if (approve is not null)
            {
                bool yes;
                try
                {
                    yes = await approve(target, ct).ConfigureAwait(false);
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested)
                {
                    yes = false;
                }
                if (!yes || ct.IsCancellationRequested)
                {
                    results[i] = new TargetResult(target, ct.IsCancellationRequested ? CommandResult.Failed("Cancelled.") : CommandResult.Declined());
                    progress?.Report(results[i]);
                    return;
                }
            }
            await gate.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                CommandResult result;
                try
                {
                    ct.ThrowIfCancellationRequested();
                    started?.Invoke(target);
                    result = await SharePointClient.RetryAsync(() => command.Run(context, target), ct).ConfigureAwait(false);
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested)
                {
                    result = CommandResult.Failed("Cancelled.");
                }
                catch (Exception e)
                {
                    result = CommandResult.Failed(SharePointErrors.Describe(e));
                }
                Diagnostics.Log.Write("command", $"{command.Id} on {target.Kind}: {(result.Ok ? "ok" : "failed")}{(result.Ok ? "" : $": {result.Message}")}");
                results[i] = new TargetResult(target, result);
                progress?.Report(results[i]);
            }
            finally
            {
                gate.Release();
            }
        }).ToList();

        try
        {
            await Task.WhenAll(tasks).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            // targets that never started: cancelled
        }
        for (var i = 0; i < results.Length; i++)
        {
            results[i] ??= new TargetResult(applicable[i], CommandResult.Failed("Cancelled."));
        }
        return results;
    }
}
