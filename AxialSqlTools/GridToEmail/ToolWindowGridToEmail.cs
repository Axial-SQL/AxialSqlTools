namespace AxialSqlTools
{
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.VisualStudio.Shell;

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
    [Guid("5fea9371-d34b-40d7-a46d-7d23eda03d9c")]
    public class ToolWindowGridToEmail : ToolWindowPane
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ToolWindowGridToEmail"/> class.
        /// </summary>
        public ToolWindowGridToEmail() : base(null)
        {
            this.Caption = "Axial SQL Tools | Grid to Email";

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
            // the object returned by the Content property.
            this.Content = new ToolWindowGridToEmailControl();
        }


        public void InitializeWithParameter(string FileNameLocation)
        {
            var control = this.Content as ToolWindowGridToEmailControl;
            if (control != null)
            {
                control.FullFileName.Text = FileNameLocation;
            }

            control.UpdateLabels();
        }
    }
}
