using System.Diagnostics;
using System.Net;

using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Tests;

public sealed class ThrottleTests
{
    [Fact]
    public void Throttling_halves_how_many_run_at_once_and_successes_raise_it_again()
    {
        var t = new Throttle();
        Assert.Equal(Throttle.Ceiling, t.Limit);
        t.Throttled(TimeSpan.Zero);
        Assert.Equal(3, t.Limit);
        t.Throttled(TimeSpan.Zero);
        t.Throttled(TimeSpan.Zero);
        Assert.Equal(Throttle.Floor, t.Limit);
        for (var i = 0; i < Throttle.Recovery - 1; i++)
        {
            t.Succeeded();
        }
        Assert.Equal(Throttle.Floor, t.Limit);
        t.Succeeded();
        Assert.Equal(Throttle.Floor + 1, t.Limit);
    }

    [Fact]
    public async Task No_more_than_the_limit_run_at_once()
    {
        var t = new Throttle();
        t.Throttled(TimeSpan.Zero); // 3 at once
        var places = new List<IDisposable>();
        for (var i = 0; i < 3; i++)
        {
            places.Add(await t.EnterAsync(CancellationToken.None));
        }
        var fourth = t.EnterAsync(CancellationToken.None);
        await Task.Delay(100);
        Assert.False(fourth.IsCompleted);
        Assert.Equal(3, t.InUse);
        places[0].Dispose();
        (await fourth.WaitAsync(TimeSpan.FromSeconds(5))).Dispose();
        places[1].Dispose();
        places[2].Dispose();
        Assert.Equal(0, t.InUse);
    }

    [Fact]
    public async Task Everything_waits_out_sharepoints_pause()
    {
        var t = new Throttle();
        t.Throttled(TimeSpan.FromMilliseconds(300));
        Assert.NotNull(t.PausedUntil);
        var clock = Stopwatch.StartNew();
        using (await t.EnterAsync(CancellationToken.None))
        {
            Assert.True(clock.ElapsedMilliseconds >= 250, $"{clock.ElapsedMilliseconds} ms");
        }
        Assert.Null(t.PausedUntil);
    }

    [Fact]
    public async Task Waiting_for_a_place_can_be_cancelled()
    {
        var t = new Throttle();
        t.Throttled(TimeSpan.FromMinutes(5));
        using var cts = new CancellationTokenSource(100);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => t.EnterAsync(cts.Token));
    }

    [Theory]
    [InlineData(0, 5)]
    [InlineData(1, 10)]
    [InlineData(2, 20)]
    [InlineData(3, 40)]
    [InlineData(4, 60)]
    [InlineData(9, 60)]
    public void Without_retry_after_it_backs_off(int attempt, int seconds) =>
        Assert.Equal(TimeSpan.FromSeconds(seconds), SharePointErrors.RetryAfter(new HttpRequestException("busy", null, HttpStatusCode.TooManyRequests), attempt));

    [Fact]
    public void Retry_after_is_honoured_and_only_throttling_is_retried()
    {
        Assert.Equal(TimeSpan.FromSeconds(42), SharePointErrors.RetryAfter(new ThrottledException("busy", HttpStatusCode.TooManyRequests, TimeSpan.FromSeconds(42)), 0));
        Assert.Equal(TimeSpan.FromSeconds(300), SharePointErrors.RetryAfter(new ThrottledException("busy", HttpStatusCode.ServiceUnavailable, TimeSpan.FromHours(1)), 0));
        Assert.Null(SharePointErrors.RetryAfter(new HttpRequestException("no", null, HttpStatusCode.Forbidden), 0));
        Assert.Null(SharePointErrors.RetryAfter(new InvalidOperationException("no"), 0));
    }

    [Fact]
    public async Task A_throttled_request_pauses_the_tenant_then_goes_through()
    {
        // a tenant of its own: throttles are per tenant, and shared
        var tenant = $"https://t{Guid.NewGuid():N}.sharepoint.com";
        var client = new SharePointClient(new Connections.ConnectionProfile { TenantUrl = tenant }, new CommandCatalogTests.NoTokens(), false);
        var attempts = 0;
        var result = await client.RetryAsync(() =>
        {
            if (++attempts == 1)
            {
                throw new ThrottledException("busy", HttpStatusCode.TooManyRequests, TimeSpan.FromSeconds(1));
            }
            return Task.FromResult("ok");
        }, CancellationToken.None);
        Assert.Equal("ok", result);
        Assert.Equal(2, attempts);
        Assert.Same(client.Throttle, Throttle.For(tenant + "/"));
        Assert.Equal(3, client.Throttle.Limit);
    }
}
