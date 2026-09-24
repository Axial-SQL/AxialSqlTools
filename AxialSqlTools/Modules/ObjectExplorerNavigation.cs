using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace AxialSqlTools
{
    internal interface IObjectExplorerTreeNode
    {
        string Name { get; }
        string InvariantName { get; }
        string UniqueName { get; }
        string UrnPath { get; }
        string NavigationContext { get; }
        bool IsFolder { get; }
        IList<IObjectExplorerTreeNode> LoadChildren();
    }

    internal interface IObjectExplorerNavigationHost
    {
        IObjectExplorerTreeNode FindServer();
        void Connect();
        void Select(IObjectExplorerTreeNode node);
    }

    internal static class ObjectExplorerNavigation
    {
        public static async Task NavigateAsync(IObjectExplorerNavigationHost host, ScriptObjectSelectionItem item,
            CancellationToken cancellationToken, Func<CancellationToken, Task> wait = null)
        {
            var steps = ObjectExplorerPath.GetTreeSteps(item);
            if (wait == null) wait = token => Task.Delay(50, token);
            cancellationToken.ThrowIfCancellationRequested();
            var server = host.FindServer();
            if (server == null)
            {
                host.Connect();
                while (server == null)
                {
                    await wait(cancellationToken);
                    cancellationToken.ThrowIfCancellationRequested();
                    server = host.FindServer();
                }
            }

            var current = server;
            // Cache each branch for this invocation. EnumerateChildren(false) loads it before
            // traversal; repeatedly refreshing an excluded/missing object would just hang SSMS.
            var children = new Dictionary<IObjectExplorerTreeNode, IList<IObjectExplorerTreeNode>>();
            foreach (var step in steps)
            {
                var found = await FindAsync(current, step, children, false, 0, wait, cancellationToken);
                if (found == null)
                    throw new InvalidOperationException("Object Explorer could not find " + step.DisplayName
                        + ". Refresh its parent folder, check Object Explorer filters and permissions, and try again.");
                current = found;
            }
            cancellationToken.ThrowIfCancellationRequested();
            host.Select(current);
        }

        private static async Task<IObjectExplorerTreeNode> FindAsync(IObjectExplorerTreeNode parent,
            ObjectExplorerTreeStep step, Dictionary<IObjectExplorerTreeNode, IList<IObjectExplorerTreeNode>> cache,
            bool withinSchema, int depth, Func<CancellationToken, Task> wait, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if (depth > 8) return null;
            if (!cache.TryGetValue(parent, out var children))
            {
                // Resume on the SSMS UI thread. Only SSMS may enumerate its own tree nodes.
                await wait(token);
                token.ThrowIfCancellationRequested();
                children = parent.LoadChildren();
                token.ThrowIfCancellationRequested();
                cache.Add(parent, children);
            }

            foreach (var child in children)
            {
                token.ThrowIfCancellationRequested();
                if (step.Matches(child, withinSchema)) return child;
            }
            foreach (var child in children)
            {
                token.ThrowIfCancellationRequested();
                // Walk only the requested hierarchy, including its optional schema/system
                // folders. Never expand unrelated databases, tables or logins.
                bool schemaFolder = step.IsSchemaFolder(child);
                if (!schemaFolder && !step.IsContainer(child)) continue;
                var found = await FindAsync(child, step, cache, withinSchema || schemaFolder,
                    depth + 1, wait, token);
                if (found != null) return found;
            }
            return null;
        }
    }
}
