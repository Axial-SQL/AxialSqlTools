using AxialSqlTools;
using System;
using System.Threading;
using System.Threading.Tasks;

internal static class QueryTemplateUiBatcherTests
{
    internal static async Task FastBatches()
    {
        int checks = 0;
        int gates = 0;
        var batcher = new QueryTemplateUiBatcher(() => { checks++; return true; },
            token => { gates++; return Task.CompletedTask; }, () => 0);
        for (int operation = 1; operation <= 512; operation++)
        {
            var checkpoint = batcher.CheckpointAsync(CancellationToken.None);
            Assert(checkpoint.IsCompleted, "A synchronous ready checkpoint suspended execution.");
            await checkpoint;
            Assert(gates == (operation - 1) / 32, "The batch did not contain 32 operations.");
        }
        Assert(checks == 512, "Shell readiness was not checked at every normal checkpoint.");
        Assert(gates == 15, "512 fast operations should require 15 intervening yields and no trailing yield.");
        Assert(batcher.YieldCount == 15 && batcher.WaitMilliseconds == 0,
            "Fast checkpoints were counted as gate waits.");
    }

    internal static async Task TimeBudget()
    {
        long elapsed = 100;
        int gates = 0;
        var batcher = new QueryTemplateUiBatcher(() => true,
            token => { gates++; elapsed += 1000; return Task.CompletedTask; }, () => elapsed);
        await batcher.CheckpointAsync(CancellationToken.None);
        elapsed += 15;
        await batcher.CheckpointAsync(CancellationToken.None);
        Assert(gates == 0, "The batch yielded before its time budget elapsed.");
        elapsed++;
        await batcher.CheckpointAsync(CancellationToken.None);
        Assert(gates == 1, "A slow operation did not cause a yield at the next checkpoint.");
        await batcher.CheckpointAsync(CancellationToken.None);
        Assert(gates == 1, "Time spent awaiting the gate counted against the new batch.");
        elapsed += 16;
        await batcher.CheckpointAsync(CancellationToken.None);
        Assert(gates == 2, "The second batch did not enforce its time budget.");
        Assert(batcher.YieldCount == 2 && batcher.WaitMilliseconds == 2000,
            "The batcher did not record actual gate wait time separately from work time.");
    }

    internal static async Task OperationLimit()
    {
        int gates = 0;
        var batcher = new QueryTemplateUiBatcher(() => true,
            token => { gates++; return Task.CompletedTask; }, () => 0, maxOperations: 3);
        for (int operation = 1; operation <= 7; operation++)
        {
            await batcher.CheckpointAsync(CancellationToken.None);
            Assert(gates == (operation - 1) / 3, "The configured operation cap was not enforced.");
        }
    }

    internal static async Task Cancellation()
    {
        int checks = 0;
        int gates = 0;
        var batcher = new QueryTemplateUiBatcher(() => { checks++; return true; },
            token => { gates++; return Task.CompletedTask; }, () => 0);
        using (var cancellation = new CancellationTokenSource())
        {
            cancellation.Cancel();
            await ExpectCancellation(() => batcher.CheckpointAsync(cancellation.Token));
            Assert(checks == 0 && gates == 0, "Cancellation allowed UI access.");
        }

        using (var cancellation = new CancellationTokenSource())
        {
            var release = NewCompletion();
            var waitingBatcher = new QueryTemplateUiBatcher(() => true,
                token => release.Task, () => 0);
            Task checkpoint = waitingBatcher.CheckpointAsync(cancellation.Token, forceYield: true);
            Assert(!checkpoint.IsCompleted, "The gate did not suspend the checkpoint.");
            cancellation.Cancel();
            release.SetResult(true);
            await ExpectCancellation(() => checkpoint);
        }
    }

    internal static async Task FailedGateRetries()
    {
        int gates = 0;
        long elapsed = 0;
        var failure = new InvalidOperationException("UI gate failed");
        var batcher = new QueryTemplateUiBatcher(() => true,
            token => { gates++; elapsed += 10; if (gates == 1) throw failure; return Task.CompletedTask; }, () => elapsed);
        try
        {
            await batcher.CheckpointAsync(CancellationToken.None, forceYield: true);
            throw new Exception("A failed gate was swallowed.");
        }
        catch (InvalidOperationException ex)
        {
            Assert(ReferenceEquals(failure, ex), "The gate failure was replaced.");
        }
        Assert(batcher.YieldCount == 1 && batcher.WaitMilliseconds == 10,
            "A failed gate was not included in wait diagnostics.");
        await batcher.CheckpointAsync(CancellationToken.None);
        Assert(gates == 2, "A failed gate incorrectly authorized the next operation.");
        await batcher.CheckpointAsync(CancellationToken.None);
        Assert(gates == 2, "A successful retry did not start a fresh batch.");
        Assert(batcher.YieldCount == 2 && batcher.WaitMilliseconds == 20,
            "A successful retry did not accumulate wait diagnostics.");
    }

    internal static async Task ModalCheckpoint()
    {
        bool ready = true;
        int gates = 0;
        int waits = 0;
        int accesses = 0;
        var release = NewCompletion();
        var batcher = new QueryTemplateUiBatcher(() => ready, token =>
        {
            gates++;
            return QueryTemplateUiGate.WaitAsync(() => Task.CompletedTask, () => ready,
                retryToken => { waits++; return release.Task; }, token);
        }, () => 0);
        await batcher.CheckpointAsync(CancellationToken.None);
        ready = false;
        var access = AccessAfterCheckpoint(batcher.CheckpointAsync(CancellationToken.None), () => accesses++);
        Assert(gates == 1 && waits == 1 && accesses == 0 && !access.IsCompleted,
            "A modal dialog within a batch did not suspend menu access.");
        ready = true;
        release.SetResult(true);
        await access;
        Assert(accesses == 1, "Menu access did not resume after the modal dialog closed.");
    }

    internal static async Task ForcedStageBoundary()
    {
        bool onUiThread = true;
        int gates = 0;
        var batcher = new QueryTemplateUiBatcher(() =>
        {
            Assert(onUiThread, "Shell state was checked before the forced gate restored UI access.");
            return true;
        }, token => { gates++; onUiThread = true; return Task.CompletedTask; }, () => 0);
        await batcher.CheckpointAsync(CancellationToken.None);
        // Background discovery completed within the count and time budgets.
        onUiThread = false;
        await batcher.CheckpointAsync(CancellationToken.None, forceYield: true);
        Assert(gates == 1 && onUiThread, "A stage boundary reused an earlier UI authorization.");
        await batcher.CheckpointAsync(CancellationToken.None);
        Assert(gates == 1, "The forced stage boundary did not reset the batch.");
    }

    private static async Task AccessAfterCheckpoint(Task checkpoint, Action access)
    {
        await checkpoint;
        access();
    }

    private static async Task ExpectCancellation(Func<Task> action)
    {
        try { await action(); }
        catch (OperationCanceledException) { return; }
        throw new Exception("Expected cancellation.");
    }

    private static TaskCompletionSource<bool> NewCompletion()
    {
        return new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }
}
