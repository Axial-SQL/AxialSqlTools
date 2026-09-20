using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;

namespace AxialSqlTools.PivotGrid
{
    [Guid("2b1b0805-69b1-4230-a9dc-22ae88772b1e")]
    public sealed class PivotGridWindow : ToolWindowPane
    {
        public PivotGridWindow() : base(null)
        {
            Caption = "Pivot Grid";
            Content = new PivotGridWindowControl();
        }

        protected override void OnClose()
        {
            (Content as PivotGridWindowControl)?.Dispose();
            base.OnClose();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) (Content as PivotGridWindowControl)?.Dispose();
            base.Dispose(disposing);
        }
    }
}
