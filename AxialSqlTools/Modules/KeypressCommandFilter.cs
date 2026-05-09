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
    public class KeypressCommandFilter : IOleCommandTarget
    {
        private IOleCommandTarget nextCommandTarget;
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

            textView.GetCaretPos(out int iLine, out int iIndex);

            if (!TryGetSnippetReplacement(lineText, iIndex, out int snippetStartIndex, out int snippetEndIndex, out string newText))
                return;

            int iTargetLength = newText.Length;

            TextSpan[] pChangedSpan = new TextSpan[] { };

            IntPtr pNewText = Marshal.StringToHGlobalUni(newText);

            try
            {
                textLines.ReplaceLines(iLine, snippetStartIndex, iLine, snippetEndIndex, pNewText, iTargetLength, pChangedSpan);
            }
            finally
            {
                Marshal.FreeHGlobal(pNewText);
            }

        }

        private bool TryGetSnippetReplacement(string lineText, int caretIndex, out int snippetStartIndex, out int snippetEndIndex, out string newText)
        {
            snippetStartIndex = -1;
            snippetEndIndex = -1;
            newText = null;

            if (string.IsNullOrEmpty(lineText) || package?.globalSnippets == null || package.globalSnippets.Count == 0)
                return false;

            int searchEndIndex = Math.Min(Math.Max(caretIndex, 0), lineText.Length);

            while (searchEndIndex > 0 && char.IsWhiteSpace(lineText[searchEndIndex - 1]))
            {
                searchEndIndex--;
            }

            if (searchEndIndex == 0)
                return false;

            foreach (KeyValuePair<string, string> snippet in package.globalSnippets)
            {
                string snippetName = snippet.Key;

                if (string.IsNullOrWhiteSpace(snippetName))
                    continue;

                int candidateStartIndex = searchEndIndex - snippetName.Length;

                if (candidateStartIndex < 0)
                    continue;

                if (candidateStartIndex > 0 && !char.IsWhiteSpace(lineText[candidateStartIndex - 1]))
                    continue;

                if (!string.Equals(lineText.Substring(candidateStartIndex, snippetName.Length), snippetName, StringComparison.Ordinal))
                    continue;

                if (newText != null && snippetName.Length <= snippetEndIndex - snippetStartIndex)
                    continue;

                snippetStartIndex = candidateStartIndex;
                snippetEndIndex = searchEndIndex;
                newText = snippet.Value;
            }

            return newText != null;
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
    }
}
