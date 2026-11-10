namespace ScriptlessSharePoint.SharePoint;

/// <summary>
/// One connection's manners towards SharePoint, shared by everything that
/// connection runs (every job, and the table's own reads), as Microsoft's
/// throttling guidance asks:
/// <list type="bullet">
/// <item>When SharePoint says 429 or 503, nothing new starts until its
/// Retry-After has passed: <see cref="Throttled"/> sets the pause,
/// <see cref="WhilePausedAsync"/> waits it out.</item>
/// <item>At most <see cref="Limit"/> commands run at once across all jobs
/// (<see cref="EnterAsync"/>). Throttling halves it (down to <see cref="Floor"/>);
/// a run of successes raises it again, one at a time, up to <see cref="Ceiling"/>.</item>
/// </list>
/// </summary>
public sealed class Throttle(TimeProvider? time = null)
{
    public const int Ceiling = 6;
    public const int Floor = 2;

    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, Throttle> ByTenant = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// The tenant's throttle, one per tenant in this app: SharePoint throttles
    /// per tenant and app, and a connection's clients (and sessions, made again
    /// when the connection changes) come and go.
    /// </summary>
    public static Throttle For(string tenantUrl) => ByTenant.GetOrAdd(tenantUrl.TrimEnd('/'), _ => new Throttle());

    /// <summary>Successes in a row that raise the limit by one.</summary>
    public const int Recovery = 20;

    private readonly TimeProvider clock = time ?? TimeProvider.System;
    private readonly object sync = new();
    private readonly LinkedList<TaskCompletionSource> waiting = new();
    private int inUse;
    private int successes;
    private DateTimeOffset pausedUntil;

    /// <summary>How many commands may run at once now.</summary>
    public int Limit { get; private set; } = Ceiling;

    /// <summary>How many are running.</summary>
    public int InUse
    {
        get
        {
            lock (sync)
            {
                return inUse;
            }
        }
    }

    /// <summary>When SharePoint said to wait until; null when it hasn't (or the wait is over).</summary>
    public DateTimeOffset? PausedUntil
    {
        get
        {
            lock (sync)
            {
                return pausedUntil > clock.GetUtcNow() ? pausedUntil : null;
            }
        }
    }

    /// <summary>Raised when SharePoint throttles, and when the limit changes (on any thread).</summary>
    public event Action? Changed;

    /// <summary>SharePoint said to slow down: pause everything for <paramref name="wait"/>, and run fewer at once.</summary>
    public void Throttled(TimeSpan wait)
    {
        lock (sync)
        {
            var until = clock.GetUtcNow() + wait;
            if (until > pausedUntil)
            {
                pausedUntil = until;
            }
            Limit = Math.Max(Floor, Limit / 2);
            successes = 0;
        }
        Changed?.Invoke();
    }

    /// <summary>A request went through: enough in a row lets one more run at once.</summary>
    public void Succeeded()
    {
        var raised = false;
        lock (sync)
        {
            if (Limit < Ceiling && ++successes >= Recovery)
            {
                Limit++;
                successes = 0;
                raised = true;
                ReleaseWaiting();
            }
        }
        if (raised)
        {
            Changed?.Invoke();
        }
    }

    /// <summary>Waits while SharePoint's pause lasts.</summary>
    public async Task WhilePausedAsync(CancellationToken ct)
    {
        while (true)
        {
            TimeSpan left;
            lock (sync)
            {
                left = pausedUntil - clock.GetUtcNow();
            }
            if (left <= TimeSpan.Zero)
            {
                return;
            }
            await Task.Delay(left, clock, ct).ConfigureAwait(false);
        }
    }

    /// <summary>A place among those running: waits for SharePoint's pause, then for a free place. Dispose it to give it back.</summary>
    public async Task<IDisposable> EnterAsync(CancellationToken ct)
    {
        while (true)
        {
            await WhilePausedAsync(ct).ConfigureAwait(false);
            TaskCompletionSource? turn = null;
            lock (sync)
            {
                if (inUse < Limit)
                {
                    inUse++;
                    return new Place(this);
                }
                turn = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
                waiting.AddLast(turn);
            }
            using (ct.Register(() => turn.TrySetCanceled(ct)))
            {
                try
                {
                    await turn.Task.ConfigureAwait(false);
                }
                finally
                {
                    lock (sync)
                    {
                        waiting.Remove(turn);
                    }
                }
            }
        }
    }

    private void Leave()
    {
        lock (sync)
        {
            inUse--;
            ReleaseWaiting();
        }
    }

    /// <summary>Wakes as many waiting as there are free places (they check again).</summary>
    private void ReleaseWaiting()
    {
        var free = Limit - inUse;
        for (var node = waiting.First; node is not null && free > 0; node = node.Next)
        {
            node.Value.TrySetResult();
            free--;
        }
    }

    private sealed class Place(Throttle owner) : IDisposable
    {
        private int disposed;

        public void Dispose()
        {
            if (Interlocked.Exchange(ref disposed, 1) == 0)
            {
                owner.Leave();
            }
        }
    }
}

/// <summary>SharePoint said 429 or 503 to a REST call; <see cref="RetryAfter"/> is how long it asked to wait, when it said.</summary>
public sealed class ThrottledException(string? message, System.Net.HttpStatusCode status, TimeSpan? retryAfter)
    : HttpRequestException(message, null, status)
{
    public TimeSpan? RetryAfter { get; } = retryAfter;
}
