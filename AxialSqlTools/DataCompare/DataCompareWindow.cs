using Microsoft.VisualStudio.Shell;
using System.Runtime.InteropServices;

namespace AxialSqlTools.DataCompare
{
    [Guid("7969c96b-f25a-4605-95d8-1eea7d52d6de")]
    public sealed class DataCompareWindow : ToolWindowPane
    {
        public DataCompareWindow() : base(null)
        {
            Caption = "Table Data Compare";
            Content = new DataCompareWindowControl();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) (Content as DataCompareWindowControl)?.Dispose();
            base.Dispose(disposing);
        }
    }
}
