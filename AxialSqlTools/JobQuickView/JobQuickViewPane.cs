using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Runtime.InteropServices;
using System.Windows.Controls;

namespace AxialSqlTools.JobQuickView
{
    [Guid("7ea6f256-c4e9-4e18-a1ac-93470f3f6bda")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class JobQuickViewPane : ToolWindowPane, IVsWindowFrameNotify3
    {
        private readonly ContentControl host = new ContentControl();
        private JobQuickViewWindow control;
        private JobQuickViewSelection selection;
        private bool released;
        internal event EventHandler Closed;

        public JobQuickViewPane() : base(null)
        {
            Caption = "Quick Manage";
            Content = host;
        }

        internal void Initialize(JobQuickViewSelection selectedJob)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (released || control != null) throw new InvalidOperationException("The job tab is already initialized or closed.");
            selection = selectedJob ?? throw new ArgumentNullException(nameof(selectedJob));
            control = new JobQuickViewWindow(new JobQuickViewService(selection.CreateConnection), selection.ServerName, selection.JobName);
            control.CaptionChanged += CaptionChanged;
            Caption = control.Caption;
            host.Content = control;
        }

        internal bool Matches(JobQuickViewSelection selectedJob)
            => !released && control != null && selection.HasSameConnection(selectedJob) &&
                string.Equals(control.JobName, selectedJob.JobName, StringComparison.Ordinal);

        private void CaptionChanged(object sender, EventArgs e) { Caption = control.Caption; }

        public override void OnToolWindowCreated()
        {
            base.OnToolWindowCreated();
            ThreadHelper.ThrowIfNotOnUIThread();
            // Register before showing so tab-close and Close All both honor draft protection.
            ErrorHandler.ThrowOnFailure(((IVsWindowFrame)Frame).SetProperty((int)__VSFPROPID.VSFPROPID_ViewHelper, this));
        }

        public int OnClose(ref uint saveOptions)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (control == null || control.CanClose()) return VSConstants.S_OK;
            ((IVsWindowFrame)Frame).Show();
            return VSConstants.E_ABORT;
        }

        public int OnShow(int show) => VSConstants.S_OK;
        public int OnMove(int x, int y, int width, int height) => VSConstants.S_OK;
        public int OnSize(int x, int y, int width, int height) => VSConstants.S_OK;
        public int OnDockableChange(int dockable, int x, int y, int width, int height) => VSConstants.S_OK;

        protected override void OnClose()
        {
            Release();
            base.OnClose();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) Release();
            base.Dispose(disposing);
        }

        private void Release()
        {
            if (released) return;
            released = true;
            if (control != null)
            {
                control.CaptionChanged -= CaptionChanged;
                control.Dispose();
                host.Content = null;
                control = null;
            }
            selection = null;
            Closed?.Invoke(this, EventArgs.Empty);
            Closed = null;
        }
    }
}
