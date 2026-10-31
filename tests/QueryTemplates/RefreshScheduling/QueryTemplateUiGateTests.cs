using AxialSqlTools;
using System;
using System.Threading;
using System.Threading.Tasks;

internal static class QueryTemplateUiGateTests
{
    internal static async Task Ready()
    {
        int yields = 0;
        int checks = 0;
        int waits = 0;
        await QueryTemplateUiGate.WaitAsync(
            () => { yields++; return Task.CompletedTask; },
            () => { checks++; return true; },
            token => { waits++; return Task.CompletedTask; }, CancellationToken.None);
        Assert(yields == 1 && checks == 1 && waits == 0, "Ready UI was retried or accessed before yielding.");
    }

    internal static async Task ModalWait()
    {
        bool ready = false;
        int checks = 0;
        int waits = 0;
        int accesses = 0;
        var release = NewCompletion();
        var gate = QueryTemplateUiGate.WaitAsync(() => Task.CompletedTask,
            () => { checks++; return ready; },
            token => { waits++; return release.Task; }, CancellationToken.None);
        var operation = AccessAfterGate(gate, () => accesses++);

        Assert(checks == 1 && waits == 1, "A modal UI was polled without waiting.");
        Assert(!operation.IsCompleted && accesses == 0, "Menu access occurred while the UI was modal.");
        ready = true;
        release.SetResult(true);
        await operation;
        Assert(checks == 2 && waits == 1 && accesses == 1, "The modal wait did not resume exactly once.");
    }

    internal static async Task RecheckAfterDiscovery()
    {
        bool ready = true;
        int accesses = 0;
        int waits = 0;
        var release = NewCompletion();
        Func<Task> yield = () => Task.CompletedTask;
        Func<bool> check = () => ready;
        Func<CancellationToken, Task> wait = token => { waits++; return release.Task; };

        await QueryTemplateUiGate.WaitAsync(yield, check, wait, CancellationToken.None);
        // A modal dialog opens while filesystem discovery is in progress.
        ready = false;
        var operation = AccessAfterGate(QueryTemplateUiGate.WaitAsync(yield, check, wait, CancellationToken.None),
            () => accesses++);
        Assert(!operation.IsCompleted && waits == 1 && accesses == 0,
            "An earlier ready state incorrectly authorized later menu access.");
        ready = true;
        release.SetResult(true);
        await operation;
        Assert(accesses == 1, "Menu access did not resume after the later modal dialog closed.");
    }

    internal static async Task Cancellation()
    {
        using (var cancelled = new CancellationTokenSource())
        {
            cancelled.Cancel();
            int calls = 0;
            var gate = QueryTemplateUiGate.WaitAsync(
                () => { calls++; return Task.CompletedTask; },
                () => { calls++; return true; },
                token => { calls++; return Task.CompletedTask; }, cancelled.Token);
            await ExpectCancellation(gate);
            Assert(calls == 0, "A cancelled request entered the UI gate.");
        }

        using (var cancellation = new CancellationTokenSource())
        {
            int checks = 0;
            int accesses = 0;
            var gate = QueryTemplateUiGate.WaitAsync(() => Task.CompletedTask,
                () => { checks++; return false; },
                token => Task.Delay(Timeout.Infinite, token), cancellation.Token);
            var operation = AccessAfterGate(gate, () => accesses++);
            Assert(checks == 1 && !operation.IsCompleted, "The gate did not suspend its modal wait.");
            cancellation.Cancel();
            await ExpectCancellation(operation);
            Assert(checks == 1 && accesses == 0, "Cancellation allowed another UI check or menu access.");
        }

        using (var cancellation = new CancellationTokenSource())
        {
            var yielded = NewCompletion();
            int checks = 0;
            var gate = QueryTemplateUiGate.WaitAsync(() => yielded.Task,
                () => { checks++; return true; }, token => Task.CompletedTask, cancellation.Token);
            cancellation.Cancel();
            yielded.SetResult(true);
            await ExpectCancellation(gate);
            Assert(checks == 0, "Cancellation during the UI yield allowed a state check.");
        }
    }

    internal static async Task StateCheckFailure()
    {
        int waits = 0;
        var expected = new InvalidOperationException("shell state unavailable");
        var gate = QueryTemplateUiGate.WaitAsync(() => Task.CompletedTask,
            () => { throw expected; },
            token => { waits++; return Task.CompletedTask; }, CancellationToken.None);
        try
        {
            await gate;
            throw new Exception("A failed shell state check was suppressed.");
        }
        catch (InvalidOperationException ex)
        {
            Assert(ReferenceEquals(expected, ex), "The state-check exception was replaced.");
        }
        Assert(waits == 0, "A failed shell state check caused a retry loop.");
    }

    private static async Task AccessAfterGate(Task gate, Action access)
    {
        await gate;
        access();
    }

    private static async Task ExpectCancellation(Task task)
    {
        try { await task; }
        catch (OperationCanceledException) { return; }
        throw new InvalidOperationException("Expected UI gate cancellation.");
    }

    private static TaskCompletionSource<bool> NewCompletion()
    {
        return new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }
}
