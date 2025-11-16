using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TextManager.Interop;

namespace AxialSqlTools
{
    internal sealed class TextManagerEventsSink : IVsTextManagerEvents
    {
        private readonly AxialSqlToolsPackage _package;

        public TextManagerEventsSink(AxialSqlToolsPackage package)
        {
            _package = package;
        }

        public int OnRegisterView(IVsTextView pView)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _package.TryAttachSnippetFilter(pView);
            return VSConstants.S_OK;
        }

        public int OnUnregisterView(IVsTextView pView)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _package.ForgetCommandFilter(pView);
            return VSConstants.S_OK;
        }

        public int OnReplaceAllInFilesBegin()
        {
            return VSConstants.S_OK;
        }

        public int OnReplaceAllInFilesEnd()
        {
            return VSConstants.S_OK;
        }

        public int OnRegisterMarkerType(int iMarkerType)
        {
            return VSConstants.S_OK;
        }

        public int OnUserPreferencesChanged(VIEWPREFERENCES[] pViewPrefs, FRAMEPREFERENCES[] pFramePrefs, LANGPREFERENCES[] pLangPrefs, FONTCOLORPREFERENCES[] pColorPrefs)
        {
            return VSConstants.S_OK;
        }
    }
}
