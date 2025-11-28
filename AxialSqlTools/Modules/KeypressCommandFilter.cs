using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace AxialSqlTools
{
    public class KeypressCommandFilter : IOleCommandTarget, IVsTextViewFilter
    {
        private IOleCommandTarget nextCommandTarget;
        private IVsTextViewFilter nextTextViewFilter;
        private IVsTextView textView;
        private AxialSqlToolsPackage package;

        public KeypressCommandFilter(AxialSqlToolsPackage package, IVsTextView textView)
        {
            this.package = package;
            this.textView = textView;
        }

        public void AddToChain()
        {
            // Adds this filter into the command chain
            if (textView != null && textView.AddCommandFilter(this, out nextCommandTarget) != VSConstants.S_OK)
            {
                throw new Exception("Failed to add command filter");
            }

            nextTextViewFilter = nextCommandTarget as IVsTextViewFilter;
        }

        public int Exec(ref Guid cmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            if (cmdGroup == VSConstants.VSStd2K && nCmdID == (uint)VSConstants.VSStd2KCmdID.RETURN && ShouldProcessReturnKey())
            {
                // Logic to handle the RETURN key press
                // Example: get the current line's text from the editor
                string currentLineText = GetCurrentLineText();

                if (currentLineText.Length > 0)
                {
                    ReplaceSnippetText(currentLineText);
                }
            }

            // Pass along the command so that other command handlers can process it
            return nextCommandTarget?.Exec(ref cmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut) ?? VSConstants.S_OK;
        }

        private bool ShouldProcessReturnKey()
        {
            var snippetSettings = SettingsManager.GetSnippetSettings();

            if (!snippetSettings.useSnippets)
            {
                return false;
            }

            switch (snippetSettings.replaceKey)
            {
                case SettingsManager.SnippetReplaceKey.Enter:
                    return true;
                case SettingsManager.SnippetReplaceKey.ShiftEnter:
                    return (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;
                default:
                    return false;
            }
        }

        private string GetCurrentLineText()
        {
            // Obtain the IVsTextLines interface from the IVsTextView
            if (textView.GetBuffer(out IVsTextLines textLines) != VSConstants.S_OK)
                return ""; // Early return if we fail to get the buffer

            // Get the caret position in the text view
            textView.GetCaretPos(out int iLine, out int iIndex);

            textLines.GetLengthOfLine(iLine, out int iLength);

            textLines.GetLineText(iLine, 0, iLine, iLength, out string lineText);

            return lineText;
        }

        private void ReplaceSnippetText(string lineText)
        {

            if (textView.GetBuffer(out IVsTextLines textLines) != VSConstants.S_OK)
                return;

            if (!package.globalSnippets.TryGetValue(lineText.Trim(), out string newText))
                return;        

            textView.GetCaretPos(out int iLine, out int iIndex);

            int iSourceLength = lineText.Length;
            int iTargetLength = newText.Length;

            TextSpan[] pChangedSpan = new TextSpan[] { };

            IntPtr pNewText = Marshal.StringToHGlobalUni(newText);

            try
            {
                textLines.ReplaceLines(iLine, 0, iLine, iSourceLength, pNewText, iTargetLength, pChangedSpan);
            }
            finally
            {
                Marshal.FreeHGlobal(pNewText);
            }

        }

        public int QueryStatus(ref Guid cmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            if (cmdGroup == VSConstants.VSStd2K)
            {
                for (int i = 0; i < prgCmds.Length; i++)
                {
                    if (prgCmds[i].cmdID == (uint)VSConstants.VSStd2KCmdID.RETURN)
                    {
                        prgCmds[i].cmdf = (uint)(OLECMDF.OLECMDF_ENABLED | OLECMDF.OLECMDF_SUPPORTED);
                        return VSConstants.S_OK;
                    }
                }
            }

            return nextCommandTarget?.QueryStatus(ref cmdGroup, cCmds, prgCmds, pCmdText) ?? VSConstants.S_OK;
        }

        public int GetWordExtent(int iLine, int iIdx, uint dwFlags, TextSpan[] pSpan)
        {
            if (pSpan == null || pSpan.Length == 0)
            {
                return nextTextViewFilter?.GetWordExtent(iLine, iIdx, dwFlags, pSpan) ?? VSConstants.E_FAIL;
            }

            bool isDoubleClick = (dwFlags & (uint)WORDEXTFLAGS.WORDEXT_DBLCLICK) != 0;

            if (isDoubleClick && textView.GetBuffer(out IVsTextLines textLines) == VSConstants.S_OK)
            {
                textLines.GetLengthOfLine(iLine, out int iLength);
                textLines.GetLineText(iLine, 0, iLine, iLength, out string lineText);

                if (TryGetQuotedStringSpan(lineText, iIdx, out int startIndex, out int endIndex))
                {
                    pSpan[0].iStartLine = iLine;
                    pSpan[0].iStartIndex = startIndex;
                    pSpan[0].iEndLine = iLine;
                    pSpan[0].iEndIndex = endIndex + 1; // TextSpan end index is exclusive

                    return VSConstants.S_OK;
                }
            }

            return nextTextViewFilter?.GetWordExtent(iLine, iIdx, dwFlags, pSpan) ?? VSConstants.E_FAIL;
        }

        public int GetPairExtents(int iLine, int iIndex, TextSpan[] pts)
        {
            if (nextTextViewFilter != null)
            {
                return nextTextViewFilter.GetPairExtents(iLine, iIndex, pts);
            }

            return VSConstants.E_NOTIMPL;
        }

        public int GetDataTipText(TextSpan[] pSpan, out string pbstrText)
        {
            pbstrText = null;

            if (nextTextViewFilter != null)
            {
                return nextTextViewFilter.GetDataTipText(pSpan, out pbstrText);
            }

            return VSConstants.E_NOTIMPL;
        }

        private static bool TryGetQuotedStringSpan(string lineText, int position, out int startIndex, out int endIndex)
        {
            startIndex = -1;
            endIndex = -1;

            if (string.IsNullOrEmpty(lineText) || position < 0 || position > lineText.Length)
            {
                return false;
            }

            int searchIndex = 0;

            while (searchIndex < lineText.Length)
            {
                int openingQuote = lineText.IndexOf('\'', searchIndex);

                if (openingQuote == -1)
                {
                    break;
                }

                int closingQuote = FindClosingQuote(lineText, openingQuote + 1);

                if (closingQuote == -1)
                {
                    break;
                }

                if (position >= openingQuote && position <= closingQuote)
                {
                    startIndex = openingQuote;
                    endIndex = closingQuote;

                    if (openingQuote > 0 && (lineText[openingQuote - 1] == 'N' || lineText[openingQuote - 1] == 'n'))
                    {
                        startIndex--;
                    }

                    return true;
                }

                searchIndex = closingQuote + 1;
            }

            return false;
        }

        private static int FindClosingQuote(string lineText, int startIndex)
        {
            int index = startIndex;

            while (index < lineText.Length)
            {
                int quoteIndex = lineText.IndexOf('\'', index);

                if (quoteIndex == -1)
                {
                    return -1;
                }

                if (quoteIndex + 1 < lineText.Length && lineText[quoteIndex + 1] == '\'')
                {
                    index = quoteIndex + 2; // Skip escaped quotes
                    continue;
                }

                return quoteIndex;
            }

            return -1;
        }
    }
}
