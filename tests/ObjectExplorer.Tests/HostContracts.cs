// Portable contracts for exercising the production host adapter on non-Windows machines.
// These are NOT SSMS/WinForms implementations. Compile validation uses the real SSMS 22
// interfaces separately; an installed SSMS smoke test is still required.
using System;
using System.Collections.Generic;

namespace Microsoft.VisualStudio.Shell
{
    internal static class ThreadHelper { public static void ThrowIfNotOnUIThread() { } }
}
namespace Microsoft.SqlServer.Management.Common
{
    internal sealed class SqlConnectionInfo
    {
        public enum AuthenticationMethod { NotSpecified, SqlPassword, ActiveDirectoryInteractive }
        public string ServerName { get; set; }
        public string UserName { get; set; }
        public bool UseIntegratedSecurity { get; set; }
        public AuthenticationMethod Authentication { get; set; }
    }
}
namespace Microsoft.SqlServer.Management.Smo.RegSvrEnum
{
    internal sealed class UIConnectionInfo
    {
        public string ServerName { get; set; }
        public string UserName { get; set; }
    }
}
namespace Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer
{
    internal interface INodeInformation
    {
        Microsoft.SqlServer.Management.Common.SqlConnectionInfo Connection { get; }
        string Name { get; }
        string InvariantName { get; }
        string UrnPath { get; }
        string NavigationContext { get; }
        object this[string name] { get; }
    }
    internal interface IObjectExplorerService { void ConnectToServer(object connection); }
}
namespace System.Windows.Forms
{
    internal class TreeNode
    {
        public string Text { get; set; }
        public object Tag { get; set; }
        public TreeNode Parent { get; set; }
        public TreeView TreeView { get; set; }
        public List<TreeNode> Nodes { get; } = new List<TreeNode>();
        public bool Expanded, Visible;
        public void Expand() { Expanded = true; }
        public void EnsureVisible() { Visible = true; }
    }
    internal sealed class TreeView
    {
        public List<TreeNode> Nodes { get; } = new List<TreeNode>();
        public TreeNode SelectedNode { get; set; }
        public bool Focused;
        public bool Focus() { Focused = true; return true; }
    }
}
