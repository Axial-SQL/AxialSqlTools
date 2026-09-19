using System;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace AxialSqlTools
{
    internal static class ToolWindowDisplay
    {
        internal static void ShowAsDocument(ToolWindowPane window)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!(window?.Frame is IVsWindowFrame frame))
                throw new NotSupportedException("Cannot create tool window");

            // Set the mode before showing, including when reopening a floating frame.
            ErrorHandler.ThrowOnFailure(frame.SetProperty(
                (int)__VSFPROPID.VSFPROPID_FrameMode, (int)VSFRAMEMODE.VSFM_MdiChild));
            ErrorHandler.ThrowOnFailure(frame.Show());
        }
    }
}
