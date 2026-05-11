using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace AxialSqlTools
{
    [Guid("51633be1-78c5-4b20-b317-2123ca70d3bd")]
    public class StatisticsSummaryWindow : ToolWindowPane, IVsWindowFrameNotify3
    {
        private const int FrameShowNoActivate = 13;

        public StatisticsSummaryWindow() : base(null)
        {
            Caption = "Statistics Summary";
            Content = new StatisticsSummaryWindowControl();
        }

        public int OnShow(int fShow)
        {
            switch (fShow)
            {
                case (int)__FRAMESHOW.FRAMESHOW_WinShown:
                case (int)__FRAMESHOW.FRAMESHOW_WinRestored:
                case FrameShowNoActivate:
                    StatisticsSummaryStore.SetWindowOpen(true);
                    break;
            }

            return VSConstants.S_OK;
        }

        public int OnMove(int x, int y, int width, int height)
        {
            return VSConstants.S_OK;
        }

        public int OnSize(int x, int y, int width, int height)
        {
            return VSConstants.S_OK;
        }

        public int OnDockableChange(int fDockable, int x, int y, int width, int height)
        {
            return VSConstants.S_OK;
        }

        public int OnClose(ref uint pgrfSaveOptions)
        {
            StatisticsSummaryStore.SetWindowOpen(false);
            AxialSqlToolsPackage.CancelStatisticsCapture(updateStore: false);
            return VSConstants.S_OK;
        }
    }
}