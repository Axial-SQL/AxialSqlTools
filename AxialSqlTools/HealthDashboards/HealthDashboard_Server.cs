namespace AxialSqlTools
{
    using System;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Microsoft.VisualStudio;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;

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
    [Guid("dd71d961-c184-4f59-b9fa-bc523472ad04")]
    public class HealthDashboard_Server : ToolWindowPane, IVsWindowFrameNotify3
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="HealthDashboard_Server"/> class.
        /// </summary>
        public HealthDashboard_Server() : base(null)
        {
            this.Caption = "Health Dashboard | Server";

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
            // the object returned by the Content property.
            this.Content = new HealthDashboard_ServerControl();
        }

        public int OnShow(int fShow)
        {
            switch (fShow)
            {
                case (int)__FRAMESHOW.FRAMESHOW_WinShown:
                case (int)__FRAMESHOW.FRAMESHOW_WinRestored:
                    // The window is being shown or restored. Handle reopening logic here.

                    HealthDashboard_ServerControl control = (HealthDashboard_ServerControl)this.Content;
                    control.StartMonitoring();

                    break;

                case 13:

                    HealthDashboard_ServerControl control2 = (HealthDashboard_ServerControl)this.Content;
                    control2.StartMonitoring();

                    break;

            }

            return VSConstants.S_OK;
        }

        public int OnMove(int i1, int i2, int i3, int i4)
        {
            // Handle the window move if needed
            return VSConstants.S_OK;
        }

        public int OnSize(int i1, int i2, int i3, int i4)
        {
            // Handle the window resize if needed
            return VSConstants.S_OK;
        }

        public int OnDockableChange(int fDockable, int i1, int i2, int i3, int i4)
        {
            // Handle changes in dockable state if needed
            return VSConstants.S_OK;
        }

        public int OnClose(ref uint pgrfSaveOptions)
        {
            // Handle the window close here
            // This is your opportunity to perform cleanup

            HealthDashboard_ServerControl control = (HealthDashboard_ServerControl)this.Content;
            control.Dispose(disposing: true);

            return VSConstants.S_OK;
        }

        protected override void OnClose()
        {
            if (this.Content is IDisposable disposableContent)
            {
                disposableContent.Dispose();
            }
            // Alternatively, perform any necessary cleanup if your UserControl does not implement IDisposable
            base.OnClose();
        }


    }
}
