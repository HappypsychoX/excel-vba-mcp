using ExcelVbaMcp.Excel;

namespace ExcelVbaMcp.Tests;

public sealed class ExcelComDispatcherTests
{
    [Fact]
    public async Task CancellationWhileQueued_CancelsWorkWithoutExecutingIt()
    {
        using var dispatcher = new ExcelComDispatcher();
        using var started = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        using var cancellation = new CancellationTokenSource();
        bool cancelledWorkExecuted = false;

        Task<int> blockingWork = dispatcher.InvokeAsync(
            () =>
            {
                started.Set();
                release.Wait();
                return 1;
            },
            CancellationToken.None);

        Assert.True(started.Wait(TimeSpan.FromSeconds(5)));
        Task<int> cancelledWork = dispatcher.InvokeAsync(
            () =>
            {
                cancelledWorkExecuted = true;
                return 2;
            },
            cancellation.Token);

        cancellation.Cancel();
        release.Set();

        Assert.Equal(1, await blockingWork);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => cancelledWork);
        Assert.False(cancelledWorkExecuted);
    }

    [Fact]
    public async Task Shutdown_RejectsNewWorkAfterDrainingTheDispatcher()
    {
        var dispatcher = new ExcelComDispatcher();
        Assert.Equal(42, await dispatcher.InvokeAsync(() => 42, CancellationToken.None));

        dispatcher.Dispose();

        Assert.Throws<ObjectDisposedException>(() =>
        {
            _ = dispatcher.InvokeAsync(() => 1, CancellationToken.None);
        });
    }

    [Fact]
    public async Task ShutdownWithBlockedComWork_StopsWaitingAtItsConfiguredBound()
    {
        using var started = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var dispatcher = new ExcelComDispatcher(TimeSpan.FromMilliseconds(100));

        Task<int> work = dispatcher.InvokeAsync(
            () =>
            {
                started.Set();
                release.Wait();
                return 7;
            },
            CancellationToken.None);
        Assert.True(started.Wait(TimeSpan.FromSeconds(5)));

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        dispatcher.Dispose();
        stopwatch.Stop();

        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(1), $"Shutdown waited {stopwatch.Elapsed}.");
        release.Set();
        Assert.Equal(7, await work);
    }
}
