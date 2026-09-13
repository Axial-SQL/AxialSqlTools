using EnvDTE;
using Microsoft.SqlServer.Management.Smo.RegSvrEnum;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace AxialSqlTools
{
    internal sealed class OpenInObjectExplorerCommand
    {
        public const int CommandId = 4135;
        public static readonly Guid CommandSet = ScriptSelectedObject.CommandSet;
        private readonly AsyncPackage package;
        private bool isRunning;

        private OpenInObjectExplorerCommand(AsyncPackage package, OleMenuCommandService commands)
        {
            this.package = package;
            var command = new OleMenuCommand(Execute, new CommandID(CommandSet, CommandId));
            command.BeforeQueryStatus += (sender, args) =>
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                var dte = Package.GetGlobalService(typeof(DTE)) as DTE;
                command.Enabled = !isRunning && dte?.ActiveDocument?.Selection is TextSelection;
            };
            commands.AddCommand(command);
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var commands = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commands != null) new OpenInObjectExplorerCommand(package, commands);
        }

        private void Execute(object sender, EventArgs args)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (isRunning) return;
            isRunning = true;
            package.JoinableTaskFactory.RunAsync(ExecuteAsync).FileAndForget("AxialSqlTools/OpenInObjectExplorer");
        }

        private async Task ExecuteAsync()
        {
            DTE dte = null;
            string status = null;
            try
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
                dte = Package.GetGlobalService(typeof(DTE)) as DTE;
                var document = dte?.ActiveDocument;
                var selection = document?.Selection as TextSelection;
                if (selection == null)
                    throw new InvalidOperationException("Open a connected SQL query window and select an object name.");

                var connection = ScriptFactoryAccess.GetCurrentConnectionInfo();
                if (connection?.ActiveConnectionInfo == null)
                    throw new InvalidOperationException("Connect the SQL query window to a server before opening an object in Object Explorer.");

                // Capture the document and connection before any await or focus change.
                // Switching tabs while metadata is loading must not change the target.
                var explorerConnection = new UIConnectionInfo(connection.ActiveConnectionInfo);
                string name = selection.Text.Trim();
                if (string.IsNullOrEmpty(name))
                {
                    var textDocument = document.Object("TextDocument") as TextDocument;
                    if (textDocument != null)
                    {
                        var start = textDocument.StartPoint.CreateEditPoint();
                        string sql = start.GetText(textDocument.EndPoint);
                        int offset = start.GetText(selection.ActivePoint).Length;
                        name = await Task.Run(() => SqlObjectName.GetNameAtCaret(sql, offset), package.DisposalToken);
                    }
                }
                if (string.IsNullOrWhiteSpace(name))
                    throw new InvalidOperationException("Select an object name or place the cursor within it. Comments and string literals are ignored.");

                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
                status = "Opening " + name.Replace("\r", " ").Replace("\n", " ") + " in Object Explorer...";
                dte.StatusBar.Text = status;
                var matches = await SqlObjectResolver.FindObjectsAsync(connection.FullConnectionString, name, package.DisposalToken);
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
                var selected = SqlObjectResolver.SelectObject(matches, "open in Object Explorer");
                if (selected == null) return;
                var steps = ObjectExplorerPath.GetRelativeSteps(selected);

                dte.ExecuteCommand("View.ObjectExplorer");
                var explorer = ServiceCache.ServiceProvider.GetService(typeof(IObjectExplorerService)) as IObjectExplorerService;
                if (explorer == null)
                    throw new InvalidOperationException("Object Explorer is unavailable. Open it from the View menu and try again.");

                using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(package.DisposalToken))
                {
                    timeout.CancelAfter(TimeSpan.FromSeconds(30));
                    try
                    {
                        var host = new SsmsObjectExplorerNavigationHost(explorer, explorerConnection, connection.FullConnectionString);
                        await ObjectExplorerNavigation.NavigateAsync(host, steps, timeout.Token);
                        if (selected.TypeDesc == "DEFAULT_CONSTRAINT")
                            dte.StatusBar.Text = "Selected column " + selected.ColumnName + " owning default constraint " + selected.ObjectName + ".";
                    }
                    catch (OperationCanceledException) when (!package.DisposalToken.IsCancellationRequested)
                    {
                        throw new InvalidOperationException("Object Explorer could not locate " + selected.DisplayName
                            + " on " + connection.ServerName + ". Refresh the database, check Object Explorer filters and connection permissions, and try again.");
                    }
                }
            }
            catch (OperationCanceledException) when (package.DisposalToken.IsCancellationRequested)
            {
                // SSMS is closing.
            }
            catch (Exception ex)
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                VsShellUtilities.ShowMessageBox(package, ex.Message, "Open in Object Explorer",
                    OLEMSGICON.OLEMSGICON_WARNING, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
            finally
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                isRunning = false;
                if (dte != null && status != null && !package.DisposalToken.IsCancellationRequested
                    && dte.StatusBar.Text == status)
                    dte.StatusBar.Text = string.Empty;
            }
        }
    }
}
