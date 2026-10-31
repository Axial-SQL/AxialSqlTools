using System;
using System.Threading;
using System.Threading.Tasks;

namespace AxialSqlTools
{
    internal static class QueryTemplateUiGate
    {
        internal static async Task WaitAsync(Func<Task> yieldToUi, Func<bool> canAccessUi,
            Func<CancellationToken, Task> waitBeforeRetry, CancellationToken cancellationToken)
        {
            if (yieldToUi == null) throw new ArgumentNullException(nameof(yieldToUi));
            if (canAccessUi == null) throw new ArgumentNullException(nameof(canAccessUi));
            if (waitBeforeRetry == null) throw new ArgumentNullException(nameof(waitBeforeRetry));

            // Waiting for a modal dialog to close does not run or repeat a
            // refresh. The supplied yield and retry wait release the UI thread.
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await yieldToUi();
                cancellationToken.ThrowIfCancellationRequested();
                if (canAccessUi()) return;
                await waitBeforeRetry(cancellationToken);
            }
        }
    }
}
