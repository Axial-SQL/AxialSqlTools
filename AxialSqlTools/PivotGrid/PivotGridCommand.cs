using EnvDTE;
using Microsoft.SqlServer.Management.UI.Grid;
using System.Collections;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Threading.Tasks;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;

namespace AxialSqlTools.PivotGrid
{
    internal sealed class PivotGridCommand
    {
        public const int CommandId = 4159;
        public static readonly Guid CommandSet = new Guid("45457e02-6dec-4a4d-ab22-c9ee126d23c5");
        private static AxialSqlToolsPackage owner;

        public static async Task InitializeAsync(AxialSqlToolsPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            owner = package;
            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService
                ?? throw new InvalidOperationException("The menu command service is unavailable.");
            commandService.AddCommand(new MenuCommand(Execute, new CommandID(CommandSet, CommandId)));
        }

        private static string GetSourceCaption(object resultsControl)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                var dte = Package.GetGlobalService(typeof(DTE)) as DTE;
                // Read the source editor connection without executing SQL or opening a new session.
                object connection = GridAccess.GetNonPublicField(dte?.ActiveWindow?.Object, "m_connection");
                if (connection == null)
                {
                    object executor = GridAccess.GetNonPublicField(resultsControl, "m_sqlExec");
                    connection = GridAccess.GetNonPublicField(executor, "m_conn");
                }
                if (string.Equals(GridAccess.GetProperty(connection, "State")?.ToString(), "Open", StringComparison.Ordinal)
                    && int.TryParse(Convert.ToString(GridAccess.GetProperty(connection, "ServerProcessId"),
                        CultureInfo.InvariantCulture), out int sessionId) && sessionId > 0)
                    return "Pivot Grid (" + sessionId.ToString(CultureInfo.InvariantCulture) + ")";
            }
            catch { /* Naming must not prevent opening a pivot. */ }

            try
            {
                // Older SSMS providers may not expose ServerProcessId. Use the numeric
                // session suffix shown in the source tab, allowing an unsaved marker.
                var dte = Package.GetGlobalService(typeof(DTE)) as DTE;
                var match = Regex.Match(dte?.ActiveWindow?.Caption ?? "", @"\(([1-9]\d*)\)\)*\s*\*?\s*$");
                if (match.Success) return "Pivot Grid (" + match.Groups[1].Value + ")";
            }
            catch { /* Keep the default title when no session identity is available. */ }
            return "Pivot Grid";
        }

        private static IGridControl FindSourceGrid(object resultsControl)
        {
            // Resolve grids from this query's results, regardless of keyboard focus.
            // IEnumerable supports both SSMS 21 CollectionBase and SSMS 22 List<T>.
            var page = GridAccess.GetNonPublicField(resultsControl, "m_gridResultsPage");
            var containers = GridAccess.GetNonPublicField(page, "m_gridContainers") as IEnumerable;
            if (containers == null) return null;
            GridControl first = null;
            foreach (var container in containers)
            {
                var grid = GridAccess.GetNonPublicField(container, "m_grid") as GridControl;
                if (grid == null || grid.IsDisposed || grid.GridStorage == null) continue;
                // Respect an explicitly focused result, even when that result is empty.
                if (grid.ContainsFocus) return grid;
                if (first == null && grid.GridStorage.NumRows() > 0) first = grid;
            }
            return first;
        }

        private static void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                // Resolve the source before showing the new tab changes the active window.
                var resultsControl = GridAccess.GetSQLResultsControl();
                if (resultsControl == null)
                {
                    MessageBox.Show("Open or activate a SQL query window with results, then choose Tools > Export Grid to Pivot Table.",
                        "Pivot Grid", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                var grid = FindSourceGrid(resultsControl);
                if (grid == null || grid.GridStorage.NumRows() == 0)
                {
                    MessageBox.Show("The active query window has no populated result grid. Run a query using Results to Grid, then try again.",
                        "Pivot Grid", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                string caption = GetSourceCaption(resultsControl);
                var window = owner.FindToolWindow(typeof(PivotGridWindow), owner.GetNextToolWindowId(), true);
                if (!(window?.Frame is IVsWindowFrame frame))
                    throw new InvalidOperationException("Cannot create the Pivot Grid window.");
                window.Caption = caption;
                ErrorHandler.ThrowOnFailure(frame.SetProperty((int)__VSFPROPID.VSFPROPID_FrameMode, VSFRAMEMODE.VSFM_MdiChild));
                ErrorHandler.ThrowOnFailure(frame.Show());
                var view = (PivotGridWindowControl)window.Content;
                owner.JoinableTaskFactory.RunAsync(() => view.LoadGridAsync(grid))
                    .FileAndForget("AxialSqlTools/PivotGrid/Load");
            }
            catch (Exception ex)
            {
                AxialSqlToolsPackage._logger?.Error(ex, "Could not open Pivot Grid");
                MessageBox.Show("Could not read results from the active query window. Wait for the query to finish and try again. If the problem continues, see the Axial SQL Tools log.",
                    "Pivot Grid", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }
}
