using Microsoft.Data.SqlClient;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Smo.RegSvrEnum;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;

namespace AxialSqlTools
{
    internal sealed class SsmsObjectExplorerNavigationHost : IObjectExplorerNavigationHost
    {
        private readonly IObjectExplorerService explorer;
        private readonly UIConnectionInfo connection;
        private readonly SqlConnectionStringBuilder sqlConnection;
        private readonly HashSet<TreeNode> expansionRequested = new HashSet<TreeNode>();

        public SsmsObjectExplorerNavigationHost(IObjectExplorerService explorer, UIConnectionInfo connection, string connectionString)
        {
            this.explorer = explorer ?? throw new ArgumentNullException(nameof(explorer));
            this.connection = connection ?? throw new ArgumentNullException(nameof(connection));
            sqlConnection = new SqlConnectionStringBuilder(connectionString);
        }

        public string FindServerContext()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            explorer.GetSelectedNodes(out int count, out INodeInformation[] selected);
            if (selected != null)
            {
                foreach (var node in selected)
                {
                    var root = node;
                    while (root?.Parent != null) root = root.Parent;
                    if (MatchesConnection(root)) return root.Context;
                }
            }

            // Inspect only connected roots. Do not enumerate other servers or shorten a
            // hostname/remove a port to find a 'close enough' connection.
            var tree = ReadProperty(explorer, "Tree") as TreeView;
            if (tree != null)
            {
                foreach (TreeNode treeNode in tree.Nodes)
                {
                    var root = GetNodeInformation(treeNode);
                    if (MatchesConnection(root)) return root.Context;
                }
            }

            foreach (string name in new[] { connection.ServerName, connection.ServerName.ToUpperInvariant(), connection.ServerName.ToLowerInvariant() })
            {
                var root = explorer.FindNode(ObjectExplorerPath.ServerUrn(name));
                if (MatchesConnection(root)) return root.Context;
            }
            return null;
        }

        public void Connect()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // Retain SSMS authentication, encryption, and other advanced options, including
            // options that would be lost by reconstructing a server-only connection.
            explorer.ConnectToServer(connection);
        }

        public bool TryExpand(string urn)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var node = FindNode(urn);
            if (node == null) return false;
            var treeNode = GetTreeNode(node);
            if (treeNode == null)
                throw new NotSupportedException("This SSMS version does not expose the Object Explorer tree needed to expand the object. Open the database in Object Explorer and try again.");
            if (!treeNode.IsExpanded && expansionRequested.Add(treeNode))
                treeNode.Expand();
            return treeNode.IsExpanded;
        }

        public bool TrySelect(string urn)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var node = FindNode(urn);
            if (node == null) return false;
            explorer.SynchronizeTree(node);
            var treeNode = GetTreeNode(node);
            if (treeNode?.TreeView != null)
            {
                treeNode.TreeView.SelectedNode = treeNode;
                treeNode.EnsureVisible();
                treeNode.TreeView.Focus();
            }
            return true;
        }

        private INodeInformation FindNode(string urn)
        {
            var node = explorer.FindNode(urn);
            // FindNode takes a URN, not a connection ID. When a server has several OE
            // connections, never silently use a node belonging to a different login.
            return MatchesConnection(node) ? node : null;
        }

        private bool MatchesConnection(INodeInformation node)
        {
            if (node?.Connection == null || string.IsNullOrEmpty(node.Context)) return false;
            var actual = node.Connection as SqlConnectionInfo;
            if (actual == null) return false;
            return ObjectExplorerPath.SameConnection(connection.ServerName, connection.UserName,
                sqlConnection.IntegratedSecurity, actual.ServerName, actual.UserName, actual.UseIntegratedSecurity);
        }

        private static INodeInformation GetNodeInformation(TreeNode node)
        {
            return (node as IServiceProvider)?.GetService(typeof(INodeInformation)) as INodeInformation
                ?? node.Tag as INodeInformation;
        }

        private static TreeNode GetTreeNode(INodeInformation node)
        {
            // SSMS 22 returns a NodeContext wrapper; older hosts expose the TreeNode itself.
            // Keep this host-specific access isolated from lookup and navigation sequencing.
            return node as TreeNode ?? ReadProperty(node, "TreeNode") as TreeNode;
        }

        private static object ReadProperty(object instance, string name)
        {
            for (Type type = instance.GetType(); type != null; type = type.BaseType)
            {
                var property = type.GetProperty(name, BindingFlags.Instance | BindingFlags.Public
                    | BindingFlags.NonPublic | BindingFlags.DeclaredOnly);
                if (property != null) return property.GetValue(instance);
            }
            return null;
        }
    }
}
