using System;
using System.Threading;
using System.Threading.Tasks;

namespace AxialSqlTools
{
    // One instance belongs to one refresh and is used serially on the UI thread.
    // The clock must be monotonic. The gate yields and waits for safe UI access.
    internal sealed class QueryTemplateUiBatcher
    {
        private readonly Func<bool> _canAccessUi;
        private readonly Func<CancellationToken, Task> _waitForUiAccess;
        private readonly Func<long> _elapsedMilliseconds;
        private readonly int _maxOperations;
        private readonly long _timeBudgetMilliseconds;
        private long _batchStartedAt;
        private int _operations;
        private bool _gateRequired;

        internal int YieldCount { get; private set; }
        internal long WaitMilliseconds { get; private set; }

        internal QueryTemplateUiBatcher(Func<bool> canAccessUi,
            Func<CancellationToken, Task> waitForUiAccess, Func<long> elapsedMilliseconds,
            int maxOperations = 32, long timeBudgetMilliseconds = 16)
        {
            _canAccessUi = canAccessUi ?? throw new ArgumentNullException(nameof(canAccessUi));
            _waitForUiAccess = waitForUiAccess ?? throw new ArgumentNullException(nameof(waitForUiAccess));
            _elapsedMilliseconds = elapsedMilliseconds ?? throw new ArgumentNullException(nameof(elapsedMilliseconds));
            if (maxOperations <= 0) throw new ArgumentOutOfRangeException(nameof(maxOperations));
            if (timeBudgetMilliseconds <= 0) throw new ArgumentOutOfRangeException(nameof(timeBudgetMilliseconds));
            _maxOperations = maxOperations;
            _timeBudgetMilliseconds = timeBudgetMilliseconds;
            _batchStartedAt = _elapsedMilliseconds();
        }

        internal Task CheckpointAsync(CancellationToken cancellationToken, bool forceYield = false)
        {
            cancellationToken.ThrowIfCancellationRequested();

            // A forced checkpoint can follow background work. Let the gate switch
            // to the UI thread before it reads shell properties in that case.
            if (forceYield || _gateRequired || !_canAccessUi()
                || _operations >= _maxOperations
                || _elapsedMilliseconds() - _batchStartedAt >= _timeBudgetMilliseconds)
            {
                _gateRequired = true;
                return WaitForNextBatchAsync(cancellationToken);
            }

            _operations++;
            return Task.CompletedTask;
        }

        private async Task WaitForNextBatchAsync(CancellationToken cancellationToken)
        {
            long waitStartedAt = _elapsedMilliseconds();
            YieldCount++;
            try
            {
                await _waitForUiAccess(cancellationToken);
            }
            finally
            {
                WaitMilliseconds += Math.Max(0, _elapsedMilliseconds() - waitStartedAt);
            }
            cancellationToken.ThrowIfCancellationRequested();
            _batchStartedAt = _elapsedMilliseconds();
            // This checkpoint authorizes the first operation of the new batch.
            // Checking before an operation avoids yielding after the final one.
            _operations = 1;
            _gateRequired = false;
        }
    }
}
