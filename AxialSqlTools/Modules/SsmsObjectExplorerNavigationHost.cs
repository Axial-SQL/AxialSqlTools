using Microsoft.Data.SqlClient;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Smo.RegSvrEnum;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace AxialSqlTools
{
    internal sealed class SsmsObjectExplorerNavigationHost : IObjectExplorerNavigationHost
    {
        private readonly IObjectExplorerService explorer;
        private readonly UIConnectionInfo connection;
        private readonly SqlConnectionStringBuilder sqlConnection;
        private readonly TreeView tree;
        private readonly Dictionary<TreeNode, Node> nodes = new Dictionary<TreeNode, Node>();
        private TreeNode serverRoot;

        public SsmsObjectExplorerNavigationHost(IObjectExplorerService explorer, UIConnectionInfo connection, string connectionString)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.explorer = explorer ?? throw new ArgumentNullException(nameof(explorer));
            this.connection = connection ?? throw new ArgumentNullException(nameof(connection));
            sqlConnection = new SqlConnectionStringBuilder(connectionString);
            tree = ObjectExplorerReflection.Read(explorer, "Tree") as TreeView;
            if (tree == null)
                throw new NotSupportedException("This SSMS version does not expose the Object Explorer tree. Open Object Explorer and try again.");
        }

        public IObjectExplorerTreeNode FindServer()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // Keep the actual root, not a server URN shared by several login connections.
            foreach (TreeNode root in tree.Nodes)
            {
                var info = GetNodeInformation(root);
                if (!(info?.Connection is SqlConnectionInfo actual)) continue;
                if (!ObjectExplorerPath.SameConnection(connection.ServerName, connection.UserName,
                    sqlConnection.IntegratedSecurity, actual.ServerName, actual.UserName, actual.UseIntegratedSecurity)) continue;
                if (!ObjectExplorerPath.SameAuthentication(sqlConnection.Authentication.ToString(), actual.Authentication.ToString())) continue;
                serverRoot = root;
                return Wrap(root);
            }
            return null;
        }

        public void Connect()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            explorer.ConnectToServer(connection);
        }

        public void Select(IObjectExplorerTreeNode node)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var target = ((Node)node).TreeNode;
            var root = target;
            while (root.Parent != null) root = root.Parent;
            if (target.TreeView != tree || root != serverRoot)
                throw new InvalidOperationException("The Object Explorer connection changed during navigation. Try again.");
            // Select the attached node directly. FindNode can return a detached NodeContext,
            // and SynchronizeTree may reconnect rather than reuse this authenticated root.
            tree.SelectedNode = target;
            target.EnsureVisible();
            if (!ObjectExplorerReflection.TryShow(explorer)) tree.Focus();
        }

        private Node Wrap(TreeNode node)
        {
            if (!nodes.TryGetValue(node, out var wrapper))
            {
                wrapper = new Node(this, node);
                nodes.Add(node, wrapper);
            }
            return wrapper;
        }

        private static INodeInformation GetNodeInformation(TreeNode node)
        {
            return (node as IServiceProvider)?.GetService(typeof(INodeInformation)) as INodeInformation
                ?? node.Tag as INodeInformation
                ?? ObjectExplorerReflection.Read(ObjectExplorerReflection.Read(node, "containedItem"), "context") as INodeInformation;
        }

        private sealed class Node : IObjectExplorerTreeNode
        {
            private readonly SsmsObjectExplorerNavigationHost host;
            private readonly INodeInformation info;
            private readonly object item;

            public Node(SsmsObjectExplorerNavigationHost host, TreeNode node)
            {
                this.host = host;
                TreeNode = node;
                info = GetNodeInformation(node);
                item = ObjectExplorerReflection.Read(node, "containedItem");
            }

            public TreeNode TreeNode { get; }
            public string Name => ObjectExplorerReflection.Read(item, "Name") as string ?? info?.Name ?? TreeNode.Text;
            public string InvariantName => info?.InvariantName;
            public string UrnPath => info?.UrnPath;
            public string NavigationContext => info?.NavigationContext;
            public string UniqueName => ObjectExplorerReflection.UniqueName(TreeNode) ?? info?["UniqueName"] as string;
            public bool IsFolder => ObjectExplorerReflection.Read(item, "IsFolder") as bool? ?? false;

            public IList<IObjectExplorerTreeNode> LoadChildren()
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                if (TreeNode.TreeView != host.tree)
                    throw new InvalidOperationException("The Object Explorer connection was closed. Reconnect and try again.");
                ObjectExplorerReflection.EnumerateChildren(TreeNode);
                TreeNode.Expand();
                var children = new List<IObjectExplorerTreeNode>();
                foreach (TreeNode child in TreeNode.Nodes) children.Add(host.Wrap(child));
                return children;
            }
        }
    }
}
