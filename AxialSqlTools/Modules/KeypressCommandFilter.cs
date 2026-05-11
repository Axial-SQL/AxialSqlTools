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
using Microsoft.IdentityModel.Tokens;
using static AxialSqlTools.AxialSqlToolsPackage;

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
            try
            {
                if (cmdGroup == VSConstants.VSStd2K && ShouldProcessKey(nCmdID))
                {
                    // Logic to handle the RETURN key press
                    // Example: get the current line's text from the editor
                    string lastWord = GetLastWord();

                    if (lastWord.Length > 0)
                    {
                        if (package.globalSnippets.TryGetValue(lastWord.ToUpper(), out string newText))
                        {
                            ReplaceSnippetText(lastWord, newText);

                            // Stop processing command
                            return VSConstants.S_OK;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, ex.StackTrace);
            }

            // Pass along the command so that other command handlers can process it
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

        private string GetLastWord()
        {
            // Obtain the IVsTextLines interface from the IVsTextView
            if (textView.GetBuffer(out IVsTextLines textLines) != VSConstants.S_OK)
                return string.Empty; // Early return if we fail to get the buffer

            // Get the caret position in the text view
            textView.GetCaretPos(out int iLine, out var iColumn);
            
            textLines.GetLineText(iLine, 0, iLine, iColumn, out string textToCursor);

            if (textToCursor.Length == 0)
                return string.Empty;

            var lastWord = textToCursor.Split(' ', '(', ')').Last();
            return lastWord;
        }

        private void ReplaceSnippetText(string lastWord, string newText)
        {
            if (textView.GetBuffer(out IVsTextLines textLines) != VSConstants.S_OK)
                return;

            textView.GetCaretPos(out int iLine, out int iColumn);

            int iSourceLength = lastWord.Length;
            var indent = iColumn - iSourceLength;
            newText = newText.Replace("\r\n", "\r\n" + new string(' ', indent));

            TextSpan[] pChangedSpan = new TextSpan[] { };
            int iTargetLength = newText.Length;
            IntPtr pNewText = Marshal.StringToHGlobalUni(newText);

            try
            {
                textLines.ReplaceLines(iLine, iColumn - iSourceLength, iLine, iColumn, pNewText, iTargetLength, pChangedSpan);
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
    }
}
