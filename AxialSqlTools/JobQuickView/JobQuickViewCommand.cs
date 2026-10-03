using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Windows.Threading;
using System.Threading.Tasks;

namespace AxialSqlTools.JobQuickView
{
    internal sealed class JobQuickViewCommand : IDisposable
    {
        public const int CommandId = 4160;
        public static readonly Guid CommandSet = ScriptSelectedObject.CommandSet;
        private readonly AsyncPackage package;
        private readonly OleMenuCommandService commands;
        private readonly OleMenuCommand command;
        private readonly DispatcherTimer attachTimer;
        private readonly Image settingsIcon;
        private IObjectExplorerService explorer;
        private TreeView tree;
        private ContextMenuStrip menu;
        private ToolStripMenuItem quickViewItem;
        private ToolStripSeparator separator;
        private bool disposed;
        private bool reportedIntegrationError;

        private JobQuickViewCommand(AsyncPackage package, OleMenuCommandService commands)
        {
            this.package = package;
            this.commands = commands;
            settingsIcon = LoadSettingsIcon();
            command = new OleMenuCommand(OpenSelectedJob, new CommandID(CommandSet, CommandId));
            command.BeforeQueryStatus += QueryStatus;
            commands.AddCommand(command);

            // Object Explorer can be created after the package. Stop polling as soon as
            // its tree exists; reattach only if SSMS disposes and recreates that tree.
            attachTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromSeconds(2)
            };
            attachTimer.Tick += AttachTimerTick;
            if (!TryAttach()) attachTimer.Start();
        }

        public static async Task<JobQuickViewCommand> InitializeAsync(AsyncPackage package)
        {
            await package.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var commands = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            return commands == null ? null : new JobQuickViewCommand(package, commands);
        }

        private void AttachTimerTick(object sender, EventArgs args)
        {
            if (disposed || package.DisposalToken.IsCancellationRequested || TryAttach()) attachTimer.Stop();
        }

        private bool TryAttach()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                explorer = ServiceCache.ServiceProvider.GetService(typeof(IObjectExplorerService)) as IObjectExplorerService;
                if (explorer == null) return false;
                var candidate = ObjectExplorerReflection.Read(explorer, "Tree") as TreeView;
                if (candidate == null || candidate.IsDisposed) return false;
                tree = candidate;
                tree.ContextMenuStripChanged += ContextMenuChanged;
                tree.Disposed += TreeDisposed;
                AttachMenu();
                return true;
            }
            catch (Exception ex)
            {
                // An incomplete attachment must not accumulate duplicate handlers on
                // the next retry (for example while SSMS replaces its menu strip).
                DetachMenu();
                if (tree != null)
                {
                    tree.ContextMenuStripChanged -= ContextMenuChanged;
                    tree.Disposed -= TreeDisposed;
                    tree = null;
                }
                // A future SSMS change must leave its own menu and the Tools fallback
                // usable, rather than breaking extension initialization.
                if (!reportedIntegrationError)
                {
                    AxialSqlToolsPackage._logger?.Warn("Quick Manage could not attach to Object Explorer ({0}). The Tools command remains available.", ex.GetType().Name);
                    reportedIntegrationError = true;
                }
                return false;
            }
        }

        private void TreeDisposed(object sender, EventArgs args)
        {
            DetachMenu();
            if (tree != null)
            {
                tree.ContextMenuStripChanged -= ContextMenuChanged;
                tree.Disposed -= TreeDisposed;
                tree = null;
            }
            if (!disposed) attachTimer.Start();
        }

        private void ContextMenuChanged(object sender, EventArgs args) => AttachMenu();

        private void AttachMenu()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (ReferenceEquals(menu, tree?.ContextMenuStrip)) return;
            DetachMenu();
            menu = tree?.ContextMenuStrip;
            if (menu == null) return;
            menu.Opening += MenuOpening;
            quickViewItem = new ToolStripMenuItem("Quick Manage")
            {
                Name = "AxialSqlTools.JobQuickView",
                ToolTipText = "Manage job steps, schedules and execution",
                Image = settingsIcon,
                ImageScaling = ToolStripItemImageScaling.SizeToFit
            };
            quickViewItem.Click += OpenContextJob;
            separator = new ToolStripSeparator();
            UpdateContextItem();
        }

        private void MenuOpening(object sender, CancelEventArgs args) => UpdateContextItem();

        private void UpdateContextItem()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                if (menu == null || menu.IsDisposed || quickViewItem == null) return;
                var selected = SelectedJob();
                var snapshot = selected == null ? null : JobQuickViewSelection.Capture(selected);
                quickViewItem.Tag = snapshot;
                quickViewItem.Visible = snapshot != null;
                separator.Visible = snapshot != null;
                // SSMS can reuse a strip and rebuild its Items collection between jobs.
                if (!menu.Items.Contains(quickViewItem)) menu.Items.Insert(0, quickViewItem);
                if (!menu.Items.Contains(separator)) menu.Items.Insert(1, separator);
            }
            catch
            {
                if (quickViewItem != null) { quickViewItem.Visible = false; quickViewItem.Tag = null; }
                if (separator != null) separator.Visible = false;
            }
        }

        private INodeInformation SelectedJob()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var service = explorer ?? ServiceCache.ServiceProvider.GetService(typeof(IObjectExplorerService)) as IObjectExplorerService;
            if (service == null) return null;
            service.GetSelectedNodes(out int count, out INodeInformation[] selected);
            if (count != 1 || selected == null || selected.Length == 0) return null;
            return JobQuickViewSelection.IsJob(selected[0]) ? selected[0] : null;
        }

        private void QueryStatus(object sender, EventArgs args)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try { command.Enabled = !disposed && SelectedJob() != null; }
            catch { command.Enabled = false; }
        }

        private void OpenSelectedJob(object sender, EventArgs args)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try { Open(JobQuickViewSelection.Capture(SelectedJob())); }
            catch (Exception ex) { ShowError(ex); }
        }

        private void OpenContextJob(object sender, EventArgs args)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                // Use the job that opened this menu, not a later tree selection.
                var selection = (sender as ToolStripMenuItem)?.Tag as JobQuickViewSelection;
                if (selection != null) Open(selection);
            }
            catch (Exception ex) { ShowError(ex); }
        }

        private void Open(JobQuickViewSelection selection)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (disposed || package.DisposalToken.IsCancellationRequested) return;
            var window = new JobQuickViewWindow(new JobQuickViewService(selection.CreateConnection), selection.ServerName, selection.JobName);
            var shell = Package.GetGlobalService(typeof(SVsUIShell)) as IVsUIShell;
            if (shell != null && shell.GetDialogOwnerHwnd(out IntPtr owner) == 0)
                new WindowInteropHelper(window).Owner = owner;
            window.Show();
        }

        private void ShowError(Exception error)
        {
            VsShellUtilities.ShowMessageBox(package, error.Message, "Quick Manage",
                OLEMSGICON.OLEMSGICON_WARNING, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        private static Image LoadSettingsIcon()
        {
            using (var stream = typeof(JobQuickViewCommand).Assembly.GetManifestResourceStream("AxialSqlTools.Resources.settings.png"))
            {
                if (stream == null) return null;
                using (var image = Image.FromStream(stream)) return new Bitmap(image);
            }
        }

        private void DetachMenu()
        {
            if (menu != null)
            {
                menu.Opening -= MenuOpening;
                if (!menu.IsDisposed)
                {
                    if (quickViewItem != null) menu.Items.Remove(quickViewItem);
                    if (separator != null) menu.Items.Remove(separator);
                }
            }
            quickViewItem?.Dispose();
            separator?.Dispose();
            quickViewItem = null;
            separator = null;
            menu = null;
        }

        public void Dispose()
        {
            if (disposed) return;
            ThreadHelper.JoinableTaskFactory.Run(async () =>
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                disposed = true;
                attachTimer.Stop();
                attachTimer.Tick -= AttachTimerTick;
                DetachMenu();
                if (tree != null)
                {
                    tree.ContextMenuStripChanged -= ContextMenuChanged;
                    tree.Disposed -= TreeDisposed;
                    tree = null;
                }
                commands.RemoveCommand(command);
                settingsIcon?.Dispose();
            });
        }
    }
}
