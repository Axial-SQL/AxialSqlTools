using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Threading.Tasks;
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

        private static void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                // Resolve the source before showing the new tab changes the active window.
                var grid = GridAccess.GetFocusGridControl();
                if (grid == null || grid.GridStorage == null || grid.GridStorage.NumRows() == 0)
                    throw new InvalidOperationException("Select a populated result grid, then choose Tools > Export Grid to Pivot Table.");
                var window = owner.FindToolWindow(typeof(PivotGridWindow), owner.GetNextToolWindowId(), true);
                if (!(window?.Frame is IVsWindowFrame frame))
                    throw new InvalidOperationException("Cannot create the Pivot Grid window.");
                ErrorHandler.ThrowOnFailure(frame.SetProperty((int)__VSFPROPID.VSFPROPID_FrameMode, VSFRAMEMODE.VSFM_MdiChild));
                ErrorHandler.ThrowOnFailure(frame.Show());
                var view = (PivotGridWindowControl)window.Content;
                owner.JoinableTaskFactory.RunAsync(() => view.LoadGridAsync(grid))
                    .FileAndForget("AxialSqlTools/PivotGrid/Load");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not open Pivot Grid. " + ex.Message, "Pivot Grid", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }
}
