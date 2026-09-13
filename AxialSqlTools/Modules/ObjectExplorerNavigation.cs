using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace AxialSqlTools
{
    internal interface IObjectExplorerNavigationHost
    {
        string FindServerContext();
        void Connect();
        bool TryExpand(string urn);
        bool TrySelect(string urn);
    }

    internal static class ObjectExplorerNavigation
    {
        public static async Task NavigateAsync(IObjectExplorerNavigationHost host, IList<string> relativeSteps,
            CancellationToken cancellationToken, Func<CancellationToken, Task> wait = null)
        {
            if (relativeSteps == null || relativeSteps.Count == 0)
                throw new ArgumentException("An object path is required.", nameof(relativeSteps));
            if (wait == null) wait = token => Task.Delay(200, token);
            cancellationToken.ThrowIfCancellationRequested();
            string server = host.FindServerContext();
            if (server == null)
            {
                host.Connect();
                while (server == null)
                {
                    await wait(cancellationToken);
                    cancellationToken.ThrowIfCancellationRequested();
                    server = host.FindServerContext();
                }
            }

            // Expand one level at a time and yield to SSMS's asynchronous metadata loader.
            // Never repeatedly restart expansion or jump straight into a cold deep path.
            await ExpandAsync(host, server, wait, cancellationToken);
            foreach (string step in relativeSteps)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string urn = server + step;
                if (step == relativeSteps[relativeSteps.Count - 1])
                {
                    while (!host.TrySelect(urn))
                    {
                        await wait(cancellationToken);
                        cancellationToken.ThrowIfCancellationRequested();
                    }
                }
                else
                    await ExpandAsync(host, urn, wait, cancellationToken);
            }
        }

        private static async Task ExpandAsync(IObjectExplorerNavigationHost host, string urn,
            Func<CancellationToken, Task> wait, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            while (!host.TryExpand(urn))
            {
                await wait(cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
            }
            await wait(cancellationToken);
        }
    }
}
