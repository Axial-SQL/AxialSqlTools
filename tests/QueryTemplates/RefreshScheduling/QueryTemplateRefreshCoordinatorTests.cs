using AxialSqlTools;
using System;
using System.Threading;
using System.Threading.Tasks;

internal static class QueryTemplateRefreshCoordinatorTests
{
    private static int _passed;

    private static async Task<int> Main()
    {
        try
        {
            await Run("Repeated requests schedule and execute only one refresh", RepeatedRequests);
            await Run("Synchronous scheduler reentry observes the published task", SynchronousSchedulerReentry);
            await Run("Refresh reentry shares the active task without a queued rerun", RefreshReentry);
            await Run("Scheduler runs outside the coordinator lock", SchedulerOutsideLock);
            await Run("Concurrent callers share one scheduled refresh", ConcurrentRequests);
            await Run("Both success and false results permit a later refresh", CompletionPermitsRefresh);
            await Run("Refresh faults complete callers and permit recovery", FaultRecovery);
            await Run("Refresh cancellation completes callers and permits recovery", CancellationRecovery);
            await Run("Scheduler exceptions complete callers and permit recovery", ScheduleFailureRecovery);
            await Run("Scheduler cancellation completes callers and permits recovery", ScheduleCancellationRecovery);
            await Run("Duplicate scheduled callbacks execute refresh only once", DuplicateScheduledCallback);
            await Run("A callback queued by a failed scheduler cannot execute later", LateCallbackAfterScheduleFailure);
            await Run("A ready UI gate yields once without retrying", QueryTemplateUiGateTests.Ready);
            await Run("A modal UI gate waits before allowing menu access", QueryTemplateUiGateTests.ModalWait);
            await Run("The UI gate rechecks modality after discovery", QueryTemplateUiGateTests.RecheckAfterDiscovery);
            await Run("UI gate cancellation before and during waits stops access", QueryTemplateUiGateTests.Cancellation);
            await Run("A failed UI state check propagates without retrying", QueryTemplateUiGateTests.StateCheckFailure);
            await Run("512 fast checkpoints batch into 15 intervening UI yields", QueryTemplateUiBatcherTests.FastBatches);
            await Run("Elapsed time ends a UI batch after slow native work", QueryTemplateUiBatcherTests.TimeBudget);
            await Run("The configured operation cap ends a UI batch", QueryTemplateUiBatcherTests.OperationLimit);
            await Run("Batch cancellation stops access before and during gates", QueryTemplateUiBatcherTests.Cancellation);
            await Run("A failed batch gate propagates and is retried", QueryTemplateUiBatcherTests.FailedGateRetries);
            await Run("Modality is checked between operations inside a UI batch", QueryTemplateUiBatcherTests.ModalCheckpoint);
            await Run("Stage boundaries force a new UI gate", QueryTemplateUiBatcherTests.ForcedStageBoundary);
            Console.WriteLine("Passed " + _passed + " Query Templates refresh scheduling tests.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }

    private static async Task Run(string name, Func<Task> test)
    {
        await WithinTimeout(test());
        _passed++;
        Console.WriteLine("PASS: " + name);
    }

    private static async Task RepeatedRequests()
    {
        var coordinator = new QueryTemplateRefreshCoordinator();
        Func<Task> callback = null;
        int scheduled = 0;
        int refreshed = 0;
        var release = NewCompletion();
        Func<Task<bool>> refresh = () => { refreshed++; return release.Task; };
        Action<Func<Task>> schedule = worker => { scheduled++; callback = worker; };

        var first = coordinator.RequestAsync(refresh, schedule);
        for (int i = 0; i < 100; i++) Same(first, coordinator.RequestAsync(refresh, schedule));
        Equal(1, scheduled);
        Equal(0, refreshed);

        var workerTask = callback();
        for (int i = 0; i < 100; i++) Same(first, coordinator.RequestAsync(refresh, schedule));
        Equal(1, refreshed);
        release.SetResult(true);
        await workerTask;
        Assert(await first, "The shared refresh result was lost.");
        Equal(1, scheduled);
        Equal(1, refreshed);
    }

    private static async Task SynchronousSchedulerReentry()
    {
        var coordinator = new QueryTemplateRefreshCoordinator();
        Task<bool> reentrant = null;
        Task workerTask = null;
        var release = NewCompletion();
        int scheduled = 0;
        Action<Func<Task>> schedule = worker =>
        {
            scheduled++;
            reentrant = coordinator.RequestAsync(() => Task.FromResult(false), UnexpectedSchedule);
            workerTask = worker();
        };

        var first = coordinator.RequestAsync(() => release.Task, schedule);
        Same(first, reentrant);
        Equal(1, scheduled);
        release.SetResult(true);
        await workerTask;
        Assert(await first, "The scheduled refresh did not complete.");
    }

    private static async Task RefreshReentry()
    {
        var coordinator = new QueryTemplateRefreshCoordinator();
        Task<bool> reentrant = null;
        Task workerTask = null;
        int refreshed = 0;
        var first = coordinator.RequestAsync(() =>
        {
            refreshed++;
            reentrant = coordinator.RequestAsync(() => Task.FromResult(false), UnexpectedSchedule);
            return Task.FromResult(true);
        }, worker => workerTask = worker());

        Same(first, reentrant);
        await workerTask;
        Assert(await first, "The original refresh result was lost.");
        Equal(1, refreshed);
    }

    private static async Task SchedulerOutsideLock()
    {
        var coordinator = new QueryTemplateRefreshCoordinator();
        Task<bool> reentrant = null;
        Func<Task> callback = null;
        var first = coordinator.RequestAsync(() => Task.FromResult(true), worker =>
        {
            // Another thread must be able to enter RequestAsync before this
            // synchronous scheduling delegate returns.
            var otherThread = Task.Run(() =>
            {
                reentrant = coordinator.RequestAsync(() => Task.FromResult(false), UnexpectedSchedule);
            });
            if (!otherThread.Wait(TimeSpan.FromSeconds(2)))
                throw new InvalidOperationException("The scheduler ran while holding the coordinator lock.");
            callback = worker;
        });

        Same(first, reentrant);
        Assert(callback != null, "Scheduling failed.");
        await callback();
        Assert(await first, "Refresh did not complete.");
    }

    private static async Task ConcurrentRequests()
    {
        var coordinator = new QueryTemplateRefreshCoordinator();
        Func<Task> callback = null;
        int scheduled = 0;
        Action<Func<Task>> schedule = worker =>
        {
            Interlocked.Increment(ref scheduled);
            callback = worker;
        };
        var first = coordinator.RequestAsync(() => Task.FromResult(true), schedule);
        var requests = new Task[32];
        for (int i = 0; i < requests.Length; i++)
            requests[i] = Task.Run(() => Same(first, coordinator.RequestAsync(() => Task.FromResult(false), schedule)));
        await Task.WhenAll(requests);
        Equal(1, scheduled);
        await callback();
        Assert(await first, "Concurrent callers replaced the original refresh.");
    }

    private static async Task CompletionPermitsRefresh()
    {
        var coordinator = new QueryTemplateRefreshCoordinator();
        Task workerTask = null;
        var first = coordinator.RequestAsync(() => Task.FromResult(true), worker => workerTask = worker());
        await workerTask;
        Assert(await first, "Expected a successful refresh.");
        var second = coordinator.RequestAsync(() => Task.FromResult(false), worker => workerTask = worker());
        Assert(!ReferenceEquals(first, second), "A completed task blocked the next request.");
        await workerTask;
        Assert(!await second, "Expected the refresh's false result to be retained.");
        await AssertRecovery(coordinator);
    }

    private static async Task FaultRecovery()
    {
        var coordinator = new QueryTemplateRefreshCoordinator();
        Task workerTask = null;
        var expected = new InvalidOperationException("refresh failed");
        var failed = coordinator.RequestAsync(() => { throw expected; }, worker => workerTask = worker());
        await workerTask;
        Same(expected, await ExpectFault<InvalidOperationException>(failed));
        Assert(failed.IsFaulted, "Refresh failure did not fault the caller task.");
        await AssertRecovery(coordinator);

        var asynchronous = NewCompletion();
        failed = coordinator.RequestAsync(() => asynchronous.Task, worker => workerTask = worker());
        asynchronous.SetException(expected);
        await workerTask;
        Same(expected, await ExpectFault<InvalidOperationException>(failed));
        await AssertRecovery(coordinator);
    }

    private static async Task CancellationRecovery()
    {
        var coordinator = new QueryTemplateRefreshCoordinator();
        Task workerTask = null;
        var cancelled = coordinator.RequestAsync(() => { throw new OperationCanceledException(); }, worker => workerTask = worker());
        await workerTask;
        await ExpectFault<OperationCanceledException>(cancelled);
        Assert(cancelled.IsCanceled, "Cancellation was converted into a fault.");
        await AssertRecovery(coordinator);

        var asynchronous = NewCompletion();
        cancelled = coordinator.RequestAsync(() => asynchronous.Task, worker => workerTask = worker());
        asynchronous.SetCanceled();
        await workerTask;
        await ExpectFault<OperationCanceledException>(cancelled);
        Assert(cancelled.IsCanceled, "Asynchronous cancellation was converted into a fault.");
        await AssertRecovery(coordinator);
    }

    private static async Task ScheduleFailureRecovery()
    {
        var coordinator = new QueryTemplateRefreshCoordinator();
        int refreshed = 0;
        var expected = new InvalidOperationException("scheduler failed");
        var failed = coordinator.RequestAsync(() => { refreshed++; return Task.FromResult(true); }, worker => { throw expected; });
        Same(expected, await ExpectFault<InvalidOperationException>(failed));
        Equal(0, refreshed);
        await AssertRecovery(coordinator);
    }

    private static async Task ScheduleCancellationRecovery()
    {
        var coordinator = new QueryTemplateRefreshCoordinator();
        var cancelled = coordinator.RequestAsync(() => Task.FromResult(true), worker => { throw new OperationCanceledException(); });
        await ExpectFault<OperationCanceledException>(cancelled);
        Assert(cancelled.IsCanceled, "Scheduler cancellation was converted into a fault.");
        await AssertRecovery(coordinator);
    }

    private static async Task DuplicateScheduledCallback()
    {
        var coordinator = new QueryTemplateRefreshCoordinator();
        Func<Task> callback = null;
        int refreshed = 0;
        var release = NewCompletion();
        var result = coordinator.RequestAsync(() => { refreshed++; return release.Task; }, worker => callback = worker);
        var running = callback();
        await callback();
        Equal(1, refreshed);
        Assert(!result.IsCompleted, "A duplicate callback completed the active request.");
        release.SetResult(true);
        await running;
        Assert(await result, "The active callback failed to publish its result.");
        await callback();
        Equal(1, refreshed);
    }

    private static async Task LateCallbackAfterScheduleFailure()
    {
        var coordinator = new QueryTemplateRefreshCoordinator();
        Func<Task> callback = null;
        int refreshed = 0;
        var failed = coordinator.RequestAsync(() => { refreshed++; return Task.FromResult(false); }, worker =>
        {
            callback = worker;
            throw new InvalidOperationException("scheduler failed after enqueueing");
        });
        await ExpectFault<InvalidOperationException>(failed);
        await AssertRecovery(coordinator);
        await callback();
        Equal(0, refreshed);
    }

    private static async Task AssertRecovery(QueryTemplateRefreshCoordinator coordinator)
    {
        Task workerTask = null;
        var recovered = coordinator.RequestAsync(() => Task.FromResult(true), worker => workerTask = worker());
        Assert(workerTask != null, "A completed request prevented recovery from being scheduled.");
        await workerTask;
        Assert(await recovered, "A later request did not recover.");
    }

    private static void UnexpectedSchedule(Func<Task> callback)
    {
        throw new InvalidOperationException("An active refresh must not schedule another worker.");
    }

    private static TaskCompletionSource<bool> NewCompletion()
    {
        return new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
    }

    private static async Task<TException> ExpectFault<TException>(Task task) where TException : Exception
    {
        try { await task; }
        catch (TException exception) { return exception; }
        throw new InvalidOperationException("Expected " + typeof(TException).Name + ".");
    }

    private static async Task WithinTimeout(Task task)
    {
        if (await Task.WhenAny(task, Task.Delay(TimeSpan.FromSeconds(5))) != task)
            throw new InvalidOperationException("A refresh scheduling test did not terminate.");
        await task;
    }

    private static void Same(object expected, object actual)
    {
        Assert(ReferenceEquals(expected, actual), "Expected the same object instance.");
    }

    private static void Equal(int expected, int actual)
    {
        Assert(expected == actual, "Expected " + expected + ", got " + actual + ".");
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }
}
