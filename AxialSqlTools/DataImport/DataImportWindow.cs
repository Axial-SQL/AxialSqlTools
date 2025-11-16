using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;

namespace AxialSqlTools
{

    [Guid("e28bf97f-0548-458e-b5d4-03033301c784")]
    public class DataImportWindow : ToolWindowPane
    {
        public DataImportWindow() : base(null)
        {
            this.Caption = "Data Import";
            this.Content = new DataImportWindowControl();
        }
    }
}
