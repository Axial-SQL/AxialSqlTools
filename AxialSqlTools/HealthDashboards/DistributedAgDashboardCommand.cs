using Microsoft.SqlServer.Management.Sdk.Sfc;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Threading.Tasks;
using System.Windows.Forms;
using ConnectionInfo = AxialSqlTools.ScriptFactoryAccess.ConnectionInfo;

namespace AxialSqlTools
{
    internal sealed class DistributedAgDashboardCommand : IDisposable
    {
        public const int CommandId = 4161;
        internal const string MenuText = "AG Health Dashboard";
        private readonly AxialSqlToolsPackage package;
        private readonly OleMenuCommandService commands;
        private readonly OleMenuCommand command;
        private readonly AvailabilityGroupContextMenu contextMenu;
        private bool disposed;

        private DistributedAgDashboardCommand(AxialSqlToolsPackage package, OleMenuCommandService commands)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.package = package;
            this.commands = commands;
            command = new OleMenuCommand(Execute, new CommandID(HealthDashboard_ServerCommand.CommandSet, CommandId));
            commands.AddCommand(command);
            contextMenu = new AvailabilityGroupContextMenu(ExecuteFromContextMenu);
        }

        public static async Task<DistributedAgDashboardCommand> InitializeAsync(AxialSqlToolsPackage package)
        {
            await package.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var commands = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            return commands == null ? null : new DistributedAgDashboardCommand(package, commands);
        }

        private void Execute(object sender, EventArgs args) => OpenDashboard(requireAvailabilityGroup: false);

        private void ExecuteFromContextMenu(object sender, EventArgs args) => OpenDashboard(requireAvailabilityGroup: true);

        private void OpenDashboard(bool requireAvailabilityGroup)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (disposed) return;
            try
            {
                // Capture the selected Object Explorer node and its authenticated connection
                // before opening the pane changes focus. Never use the active query's server.
                var node = GetSelectedNode(GetExplorer());
                if (node == null)
                    throw new InvalidOperationException("Select a connected SQL Server instance or availability group in Object Explorer, then open this dashboard.");
                bool hasGroup = TryGetAvailabilityGroup(node, out string name);
                if (requireAvailabilityGroup && !hasGroup)
                    throw new InvalidOperationException("Select an availability group in Object Explorer, then open this dashboard.");
                // The Tools command also accepts a connected instance. A null name asks
                // the dashboard to list that instance's availability groups.
                if (!hasGroup) name = null;

                ConnectionInfo connection = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer(inMaster: true);
                if (connection == null || string.IsNullOrWhiteSpace(connection.FullConnectionString))
                    throw new InvalidOperationException("The selected Object Explorer connection is unavailable. Reconnect the server and try again.");

                var window = package.FindToolWindow(typeof(DistributedAgDashboardWindow), package.GetNextToolWindowId(), true)
                    as DistributedAgDashboardWindow;
                if (window?.Frame == null)
                    throw new NotSupportedException("Cannot create the AG Health Dashboard window.");
                window.Initialize(connection, name);
                ToolWindowDisplay.ShowAsDocument(window);
            }
            catch (Exception ex)
            {
                AxialSqlToolsPackage._logger?.Error(ex, "Cannot open the AG Health Dashboard.");
                VsShellUtilities.ShowMessageBox(package, ex.Message, MenuText,
                    OLEMSGICON.OLEMSGICON_WARNING, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        private static IObjectExplorerService GetExplorer()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return ServiceCache.ServiceProvider?.GetService(typeof(IObjectExplorerService)) as IObjectExplorerService;
        }

        private static INodeInformation GetSelectedNode(IObjectExplorerService explorer)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (explorer == null) return null;
            explorer.GetSelectedNodes(out int count, out INodeInformation[] nodes);
            return count == 1 && nodes != null && nodes.Length > 0 ? nodes[0] : null;
        }

        private static bool TryGetAvailabilityGroup(INodeInformation node, out string name)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            name = null;
            if (node == null) return false;
            if (node.UrnPath != "AvailabilityGroup"
                && !(node.UrnPath?.EndsWith("/AvailabilityGroup", StringComparison.Ordinal) ?? false)) return false;
            if (string.IsNullOrWhiteSpace(node.Context)) return false;
            var urn = new Urn(node.Context);
            if (!urn.IsValidUrn() || !string.Equals(urn.Type, "AvailabilityGroup", StringComparison.Ordinal)) return false;
            name = urn.GetAttribute("Name");
            // Use the catalog URN, never the localized caption with (Primary)/(Distributed).
            // Regular and distributed AG nodes use the same dashboard.
            return !string.IsNullOrWhiteSpace(name);
        }

        public void Dispose()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (disposed) return;
            disposed = true;
            contextMenu.Dispose();
            commands.RemoveCommand(command);
        }

        /// <summary>
        /// SSMS 22 builds native Object Explorer context menus as WinForms menus.
        /// Observe menu replacement and opening rather than assuming private command GUIDs.
        /// Reflection is limited to the same Tree property used by Object Explorer navigation.
        /// </summary>
        private sealed class AvailabilityGroupContextMenu : IDisposable
        {
            private readonly EventHandler execute;
            private readonly Timer attachTimer;
            private IObjectExplorerService explorer;
            private TreeView tree;
            private ContextMenuStrip menu;
            private ToolStripMenuItem item;
            private ToolStripSeparator separator;
            private bool disposed;
            private bool reportedError;

            public AvailabilityGroupContextMenu(EventHandler execute)
            {
                this.execute = execute;
                // The package may load before Object Explorer. Retry only until the tree
                // exists; no server connection or metadata query is performed here.
                attachTimer = new Timer { Interval = 2000 };
                attachTimer.Tick += Attach;
                attachTimer.Start();
                Attach(this, EventArgs.Empty);
            }

            private void Attach(object sender, EventArgs args)
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                if (disposed || tree != null) return;
                try
                {
                    explorer = GetExplorer();
                    var candidate = ObjectExplorerReflection.Read(explorer, "Tree") as TreeView;
                    if (candidate == null || candidate.IsDisposed) return;
                    tree = candidate;
                    tree.ContextMenuStripChanged += MenuChanged;
                    tree.Disposed += TreeDisposed;
                    attachTimer.Stop();
                    MenuChanged(this, EventArgs.Empty);
                }
                catch (Exception ex)
                {
                    ReportError(ex);
                }
            }

            private void MenuChanged(object sender, EventArgs args)
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                DetachMenu();
                if (disposed || tree?.ContextMenuStrip == null || tree.ContextMenuStrip.IsDisposed) return;
                menu = tree.ContextMenuStrip;
                menu.Opening += MenuOpening;
                menu.Disposed += MenuDisposed;
                UpdateMenu();
            }

            private void MenuOpening(object sender, CancelEventArgs args) => UpdateMenu();

            private void UpdateMenu()
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                RemoveItems();
                if (disposed || menu == null || menu.IsDisposed) return;
                try
                {
                    if (!TryGetAvailabilityGroup(GetSelectedNode(explorer), out _)) return;
                    separator = new ToolStripSeparator();
                    item = new ToolStripMenuItem(MenuText) { Name = "AxialSqlTools.DistributedAgDashboard" };
                    item.Click += execute;
                    menu.Items.Add(separator);
                    menu.Items.Add(item);
                }
                catch (Exception ex)
                {
                    ReportError(ex);
                }
            }

            private void ReportError(Exception error)
            {
                if (reportedError) return;
                reportedError = true;
                AxialSqlToolsPackage._logger?.Warn(error,
                    "Cannot attach the AG Health Dashboard Object Explorer menu. The Axial SQL Tools menu remains available.");
            }

            private void RemoveItems()
            {
                if (item != null)
                {
                    item.Click -= execute;
                    menu?.Items.Remove(item);
                    item.Dispose();
                    item = null;
                }
                if (separator != null)
                {
                    menu?.Items.Remove(separator);
                    separator.Dispose();
                    separator = null;
                }
            }

            private void DetachMenu()
            {
                RemoveItems();
                if (menu == null) return;
                menu.Opening -= MenuOpening;
                menu.Disposed -= MenuDisposed;
                menu = null;
            }

            private void MenuDisposed(object sender, EventArgs args) => DetachMenu();

            private void TreeDisposed(object sender, EventArgs args)
            {
                DetachTree();
                if (!disposed) attachTimer.Start();
            }

            private void DetachTree()
            {
                DetachMenu();
                if (tree == null) return;
                tree.ContextMenuStripChanged -= MenuChanged;
                tree.Disposed -= TreeDisposed;
                tree = null;
                explorer = null;
            }

            public void Dispose()
            {
                if (disposed) return;
                disposed = true;
                attachTimer.Stop();
                attachTimer.Tick -= Attach;
                attachTimer.Dispose();
                DetachTree();
            }
        }
    }
}
