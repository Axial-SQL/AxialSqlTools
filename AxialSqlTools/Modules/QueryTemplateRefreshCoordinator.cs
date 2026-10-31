using System;
using System.Threading;
using System.Threading.Tasks;

namespace AxialSqlTools
{
    // Coalesce requests into one scheduled refresh. A request received during
    // that refresh joins it; it does not enqueue another menu rebuild.
    internal sealed class QueryTemplateRefreshCoordinator
    {
        private readonly object _sync = new object();
        private Task<bool> _active;

        internal Task<bool> RequestAsync(Func<Task<bool>> refresh, Action<Func<Task>> schedule)
        {
            if (refresh == null) throw new ArgumentNullException(nameof(refresh));
            if (schedule == null) throw new ArgumentNullException(nameof(schedule));

            TaskCompletionSource<bool> completion;
            lock (_sync)
            {
                if (_active != null && !_active.IsCompleted) return _active;
                completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                // Publish before the scheduler can invoke native code or call
                // back into RequestAsync, including a synchronous scheduler.
                _active = completion.Task;
            }

            int started = 0;
            try
            {
                // Never invoke external scheduling code while holding _sync.
                schedule(() =>
                {
                    if (Interlocked.Exchange(ref started, 1) != 0 || completion.Task.IsCompleted)
                        return Task.CompletedTask;
                    return ExecuteAsync(refresh, completion);
                });
            }
            catch (OperationCanceledException)
            {
                completion.TrySetCanceled();
            }
            catch (Exception ex)
            {
                completion.TrySetException(ex);
            }

            return completion.Task;
        }

        private static async Task ExecuteAsync(Func<Task<bool>> refresh, TaskCompletionSource<bool> completion)
        {
            try
            {
                completion.TrySetResult(await refresh());
            }
            catch (OperationCanceledException)
            {
                completion.TrySetCanceled();
            }
            catch (Exception ex)
            {
                completion.TrySetException(ex);
            }
        }
    }
}
