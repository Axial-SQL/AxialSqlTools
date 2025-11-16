using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;

namespace AxialSqlTools
{
    /// <summary>
    /// This class implements the tool window exposed by this package and hosts a user control.
    /// </summary>
    /// <remarks>
    /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane,
    /// usually implemented by the package implementer.
    /// <para>
    /// This class derives from the ToolWindowPane class provided from the MPF in order to use its
    /// implementation of the IVsUIElementPane interface.
    /// </para>
    /// </remarks>
    [Guid("2a9291bc-f953-4882-892b-693119f694d3")]
    public class DataImportWindow : ToolWindowPane
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DataImportWindow"/> class.
        /// </summary>
        public DataImportWindow() : base(null)
        {
            this.Caption = "Data Import";

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
            // the object returned by the Content property.
            this.Content = new DataImportWindowControl();
        }
    }
}
