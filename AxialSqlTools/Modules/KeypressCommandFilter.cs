using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.Runtime.InteropServices;
using System.Windows.Input;

namespace AxialSqlTools
{
    public class KeypressCommandFilter : IOleCommandTarget
    {
        private IOleCommandTarget nextCommandTarget;
        private IVsTextView textView;
        public KeypressCommandFilter(AxialSqlToolsPackage package, IVsTextView textView)
        {
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
            if (cmdGroup == VSConstants.VSStd2K && ShouldProcessKey(nCmdID))
            {
                if (TryReplaceSnippet())
                {
                    // Snippet was replaced — swallow the key so no newline/tab is inserted.
                    return VSConstants.S_OK;
                }
            }

            // Pass along the command so that other command handlers can process it.
            return nextCommandTarget?.Exec(ref cmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut) ?? VSConstants.S_OK;
        }

        private bool ShouldProcessKey(uint nCmdID)
        {
            var snippetSettings = SettingsManager.GetSnippetSettings();

            if (!snippetSettings.useSnippets)
            {
                return false;
            }

            switch (snippetSettings.replaceKey)
            {
                case SettingsManager.SnippetReplaceKey.Enter:
                    return (nCmdID == (uint)VSConstants.VSStd2KCmdID.RETURN);
                
                case SettingsManager.SnippetReplaceKey.ShiftEnter:
                    return (nCmdID == (uint)VSConstants.VSStd2KCmdID.RETURN) && 
                           (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;

                case SettingsManager.SnippetReplaceKey.Tab:
                    return (nCmdID == (uint)VSConstants.VSStd2KCmdID.TAB);

                default:
                    return false;
            }
        }

        private bool TryReplaceSnippet()
        {
            if (textView.GetBuffer(out IVsTextLines textLines) != VSConstants.S_OK)
                return false;

            textView.GetCaretPos(out int iLine, out int iColumn);

            textLines.GetLengthOfLine(iLine, out int lineLength);
            textLines.GetLineText(iLine, 0, iLine, lineLength, out string lineText);

            if (string.IsNullOrEmpty(lineText) || iColumn == 0)
                return false;

            int wordStart = iColumn;
            for (int i = iColumn - 1; i >= 0; i--)
            {
                char c = lineText[i];
                if (c == ' ' || c == '\t' || c == '(' || c == ')' || c == ',' || c == ';')
                    break;

                wordStart = i;
            }

            if (wordStart >= iColumn)
                return false;

            string word = lineText.Substring(wordStart, iColumn - wordStart).Trim();
            if (string.IsNullOrEmpty(word))
                return false;

            var dict = SnippetService.SnippetDictionary;
            if (!dict.TryGetValue(word, out SnippetItem snippet))
                return false;

            var settings = SettingsManager.GetSnippetSettings();
            var result = SnippetVariableProcessor.ProcessVariables(snippet.Body, settings.cursorMarker);
            string newText = result.ProcessedText;
            int cursorOffset = result.CursorOffset;

            IntPtr pNewText = Marshal.StringToHGlobalUni(newText);
            try
            {
                TextSpan[] pChangedSpan = new TextSpan[1];
                textLines.ReplaceLines(iLine, wordStart, iLine, iColumn, pNewText, newText.Length, pChangedSpan);
            }
            finally
            {
                Marshal.FreeHGlobal(pNewText);
            }

            SetCaretPosition(iLine, wordStart, newText, cursorOffset >= 0 ? cursorOffset : newText.Length);
            return true;
        }

        private void SetCaretPosition(int startLine, int startColumn, string text, int offset)
        {
            int targetLine = startLine;
            int targetColumn = startColumn;

            for (int i = 0; i < offset && i < text.Length; i++)
            {
                if (text[i] == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                {
                    targetLine++;
                    targetColumn = 0;
                    i++;
                }
                else if (text[i] == '\n')
                {
                    targetLine++;
                    targetColumn = 0;
                }
                else
                {
                    targetColumn++;
                }
            }

            textView.SetCaretPos(targetLine, targetColumn);
        }

        public int QueryStatus(ref Guid cmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            if (cmdGroup == VSConstants.VSStd2K)
            {
                for (int i = 0; i < prgCmds.Length; i++)
                {
                    if (prgCmds[i].cmdID == (uint)VSConstants.VSStd2KCmdID.RETURN ||
                        prgCmds[i].cmdID == (uint)VSConstants.VSStd2KCmdID.TAB)
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
