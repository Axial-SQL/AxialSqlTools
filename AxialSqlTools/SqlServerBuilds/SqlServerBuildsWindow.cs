using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

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
    [Guid("f4aa71f1-75ff-47e2-9e59-db01cdc91419")]
    public class SqlServerBuildsWindow : ToolWindowPane
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SqlServerBuildsWindow"/> class.
        /// </summary>
        public SqlServerBuildsWindow() : base(null)
        {
            this.Caption = "SQL Server Builds";

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
            // the object returned by the Content property.
            this.Content = new SqlServerBuildsWindowControl();
        }
    }
}
