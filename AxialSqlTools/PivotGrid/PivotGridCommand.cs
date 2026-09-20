using Microsoft.VisualStudio;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Threading.Tasks;
using System.Windows;

namespace AxialSqlTools.PivotGrid
{
    internal sealed class PivotGridCommand : ResultGridCommandBase
    {
        private static CommandBarButton button;
        private static AxialSqlToolsPackage owner;

        public static async Task InitializeAsync(AxialSqlToolsPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            owner = package;
            button = (CommandBarButton)GridCommandBar.Controls.Add(MsoControlType.msoControlButton,
                Type.Missing, Type.Missing, Type.Missing, true);
            button.Caption = "Pivot Grid";
            button.TooltipText = "Summarize this result grid in a new pivot window";
            button.Visible = true;
            button.Enabled = true;
            button.Click += OnClick;
        }

        private static void OnClick(CommandBarButton control, ref bool cancelDefault)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            cancelDefault = true;
            try
            {
                // Resolve the source before showing the new tab changes the active window.
                var grid = GridAccess.GetFocusGridControl();
                if (grid == null || grid.GridStorage == null || grid.GridStorage.NumRows() == 0)
                    throw new InvalidOperationException("Select a populated result grid, then choose Pivot Grid.");
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
