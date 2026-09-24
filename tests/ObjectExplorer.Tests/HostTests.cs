using AxialSqlTools;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Smo.RegSvrEnum;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

internal static partial class Program
{
    private static void HostTests()
    {
        Check("host uses attached tree under matching login; no FindNode API", () =>
        {
            var tree = new TreeView();
            var wrong = HostTree("bob");
            var right = HostTree("alice");
            Attach(tree, wrong); Attach(tree, right);
            var service = new ExplorerService(tree);
            var host = HostAdapter(service);
            ObjectExplorerNavigation.NavigateAsync(host, Item("USER_TABLE"), CancellationToken.None,
                token => Task.CompletedTask).GetAwaiter().GetResult();
            var target = right.Nodes[0].Nodes[0].Nodes[0].Nodes[0];
            Equal(target, tree.SelectedNode); True(target.Visible); True(tree.Focused);
            Equal(0, wrong.Loads); Equal(0, service.Connections);
            Equal(false, right.Async); Equal(1, right.Loads);
        });
        Check("host reuses endpoint even when root caption differs", () =>
        {
            var tree = new TreeView(); var root = HostTree("alice");
            root.Text = "unrelated display alias"; Attach(tree, root);
            True(HostAdapter(new ExplorerService(tree)).FindServer() != null);
        });
        Check("same username with different authentication does not match", () =>
        {
            var tree = new TreeView(); var root = HostTree("alice");
            root.Info.Connection.Authentication = SqlConnectionInfo.AuthenticationMethod.ActiveDirectoryInteractive;
            Attach(tree, root);
            Equal(null, HostAdapter(new ExplorerService(tree)).FindServer());
        });
        Check("legacy and explicit SQL password modes match", () =>
            True(ObjectExplorerPath.SameAuthentication("SqlPassword", "NotSpecified")));
        Check("host connects with original captured settings object", () =>
        {
            var settings = new UIConnectionInfo { ServerName = "HOST", UserName = "alice" };
            var service = new ExplorerService(new TreeView());
            var host = new SsmsObjectExplorerNavigationHost(service, settings, "Server=HOST;User ID=alice");
            host.Connect(); Equal(settings, service.ConnectedWith); Equal(1, service.Connections);
        });
        Check("host rejects unsupported tree service explicitly", () =>
            Throws<NotSupportedException>(() => HostAdapter(new MissingTreeService())));
        Check("host does not select a node detached during enumeration", () =>
        {
            var tree = new TreeView(); var root = HostTree("alice"); Attach(tree, root);
            var folder = (HostNativeNode)root.Nodes[0].Nodes[0].Nodes[0];
            folder.OnLoad = () => folder.Nodes[0].TreeView = null;
            Throws<InvalidOperationException>(() => ObjectExplorerNavigation.NavigateAsync(
                HostAdapter(new ExplorerService(tree)), Item("USER_TABLE"), CancellationToken.None,
                token => Task.CompletedTask).GetAwaiter().GetResult());
            Equal(null, tree.SelectedNode);
        });
        Check("host Show method used when available", () =>
        {
            var tree = new TreeView(); Attach(tree, HostTree("alice"));
            var service = new ShowingExplorerService(tree);
            ObjectExplorerNavigation.NavigateAsync(HostAdapter(service), Item("USER_TABLE"), CancellationToken.None,
                token => Task.CompletedTask).GetAwaiter().GetResult();
            Equal(1, service.Shows); True(tree.SelectedNode.Visible);
        });
    }

    private static SsmsObjectExplorerNavigationHost HostAdapter(IObjectExplorerService service) =>
        new SsmsObjectExplorerNavigationHost(service, new UIConnectionInfo { ServerName = "HOST", UserName = "alice" },
            "Server=HOST;User ID=alice");

    private static HostNativeNode HostTree(string login)
    {
        var target = new HostNativeNode("Orders", false, "Table", "dbo.Orders");
        var tables = new HostNativeNode("Tabellen", true, "UserTables", null, target);
        var database = new HostNativeNode("Sales", false, "Database", "Sales", tables);
        var databases = new HostNativeNode("Datenbanken", true, "Databases", null, database);
        var root = new HostNativeNode("HOST", false, "Server", "HOST", databases);
        root.Info.Connection = new SqlConnectionInfo { ServerName = "HOST", UserName = login };
        return root;
    }
    private static void Attach(TreeView tree, TreeNode node, TreeNode parent = null)
    {
        node.TreeView = tree; node.Parent = parent;
        if (parent == null) tree.Nodes.Add(node);
        foreach (var child in node.Nodes) Attach(tree, child, node);
    }
    private sealed class HostNodeInfo : INodeInformation
    {
        public SqlConnectionInfo Connection { get; set; }
        public string Name { get; set; }
        public string InvariantName { get; set; }
        public string UrnPath { get; set; }
        public string NavigationContext => null;
        public object this[string name] => null;
    }
    private class HostNativeNodeBase : TreeNode, IServiceProvider
    {
#pragma warning disable CS0414
        private readonly HostItem containedItem;
#pragma warning restore CS0414
        public HostNodeInfo Info;
        public int Loads;
        public bool Async = true;
        public Action OnLoad;
        protected HostNativeNodeBase(string name, bool folder, string id, string invariant, params HostNativeNode[] children)
        {
            Text = name;
            containedItem = new HostItem(name, folder, id);
            Info = new HostNodeInfo { Name = name, InvariantName = invariant, UrnPath = folder ? "" : "Server/" + id };
            foreach (var child in children) Nodes.Add(child);
        }
        public object GetService(Type type) => type == typeof(INodeInformation) ? Info : null;
        private void EnumerateChildren(bool async) { Async = async; Loads++; OnLoad?.Invoke(); }
    }
    private sealed class HostNativeNode : HostNativeNodeBase
    {
        public HostNativeNode(string name, bool folder, string id, string invariant, params HostNativeNode[] children)
            : base(name, folder, id, invariant, children) { }
    }
    private sealed class HostItem
    {
#pragma warning disable CS0414
        private readonly HostContext context;
#pragma warning restore CS0414
        public HostItem(string name, bool folder, string id) { Name = name; IsFolder = folder; context = new HostContext(id); }
        public string Name { get; }
        public bool IsFolder { get; }
    }
    private sealed class HostContext
    {
        private readonly string id;
        public HostContext(string id) { this.id = id; }
        public object this[string key] => key == "UniqueName" ? id : null;
    }
    private class ExplorerService : IObjectExplorerService
    {
        private TreeView Tree { get; }
        public int Connections;
        public object ConnectedWith;
        public ExplorerService(TreeView tree) { Tree = tree; }
        public void ConnectToServer(object info) { Connections++; ConnectedWith = info; }
    }
    private sealed class ShowingExplorerService : ExplorerService
    {
        public int Shows;
        public ShowingExplorerService(TreeView tree) : base(tree) { }
        public void Show() { Shows++; }
    }
    private sealed class MissingTreeService : IObjectExplorerService { public void ConnectToServer(object info) { } }
}
