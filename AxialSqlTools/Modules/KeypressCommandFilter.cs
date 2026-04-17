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
        private AxialSqlToolsPackage package;

        public KeypressCommandFilter(AxialSqlToolsPackage package, IVsTextView textView)
        {
            this.package = package;
            this.textView = textView;
        }

        public void AddToChain()
        {
            if (textView != null && textView.AddCommandFilter(this, out nextCommandTarget) != VSConstants.S_OK)
            {
                throw new Exception("Failed to add command filter");
            }
        }

        public int Exec(ref Guid cmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            if (cmdGroup == VSConstants.VSStd2K)
            {
                bool isReturn = nCmdID == (uint)VSConstants.VSStd2KCmdID.RETURN;
                bool isTab = nCmdID == (uint)VSConstants.VSStd2KCmdID.TAB;

                if ((isReturn || isTab) && ShouldProcessKey(isReturn, isTab))
                {
                    if (TryReplaceSnippet())
                    {
                        // Snippet was replaced — swallow the key so no newline/tab is inserted
                        return VSConstants.S_OK;
                    }
                }
            }

            return nextCommandTarget?.Exec(ref cmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut) ?? VSConstants.S_OK;
        }

        private bool ShouldProcessKey(bool isReturn, bool isTab)
        {
            var snippetSettings = SettingsManager.GetSnippetSettings();

            if (!snippetSettings.useSnippets)
                return false;

            switch (snippetSettings.replaceKey)
            {
                case SettingsManager.SnippetReplaceKey.Tab:
                    return isTab;
                case SettingsManager.SnippetReplaceKey.Enter:
                    return isReturn;
                case SettingsManager.SnippetReplaceKey.ShiftEnter:
                    return isReturn && (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;
                default:
                    return false;
            }
        }

        private bool TryReplaceSnippet()
        {
            if (textView.GetBuffer(out IVsTextLines textLines) != VSConstants.S_OK)
                return false;

            textView.GetCaretPos(out int iLine, out int iColumn);

            // Get the current line text
            textLines.GetLengthOfLine(iLine, out int lineLength);
            textLines.GetLineText(iLine, 0, iLine, lineLength, out string lineText);

            if (string.IsNullOrEmpty(lineText) || iColumn == 0)
                return false;

            // Extract the current word: walk backwards from cursor position
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

            // Lookup in SnippetService dictionary (case-insensitive)
            var dict = SnippetService.SnippetDictionary;
            if (!dict.TryGetValue(word, out SnippetItem snippet))
                return false;

            // Process variables and cursor marker
            var settings = SettingsManager.GetSnippetSettings();
            var result = SnippetVariableProcessor.ProcessVariables(snippet.Body, settings.cursorMarker);
            string newText = result.ProcessedText;
            int cursorOffset = result.CursorOffset;

            // Replace only the word span with the snippet body
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

            // Position cursor
            if (cursorOffset >= 0)
            {
                // Calculate absolute position from the start of the inserted text
                int absoluteOffset = cursorOffset;
                int targetLine = iLine;
                int targetColumn = wordStart;

                // Walk through the processed text to find line/column for cursorOffset
                for (int i = 0; i < absoluteOffset && i < newText.Length; i++)
                {
                    if (newText[i] == '\r' && i + 1 < newText.Length && newText[i + 1] == '\n')
                    {
                        targetLine++;
                        targetColumn = 0;
                        i++; // skip \n
                    }
                    else if (newText[i] == '\n')
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
            else
            {
                // No cursor marker — position at the end of inserted text
                int targetLine = iLine;
                int targetColumn = wordStart;

                for (int i = 0; i < newText.Length; i++)
                {
                    if (newText[i] == '\r' && i + 1 < newText.Length && newText[i + 1] == '\n')
                    {
                        targetLine++;
                        targetColumn = 0;
                        i++;
                    }
                    else if (newText[i] == '\n')
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

            return true;
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
