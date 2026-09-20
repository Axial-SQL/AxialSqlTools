using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using System.Windows.Input;
using Task = System.Threading.Tasks.Task;

namespace AxialSqlTools
{
    internal sealed class SortSelectedTextCommand
    {
        public const int AscendingCommandId = 4157;
        public const int DescendingCommandId = 4158;
        public static readonly Guid CommandSet = new Guid("45457e02-6dec-4a4d-ab22-c9ee126d23c5");
        private readonly AsyncPackage package;

        private SortSelectedTextCommand(AsyncPackage package, OleMenuCommandService commands)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            if (commands == null) throw new ArgumentNullException(nameof(commands));
            AddCommand(commands, AscendingCommandId, false);
            AddCommand(commands, DescendingCommandId, true);
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var commands = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            new SortSelectedTextCommand(package, commands);
        }

        private void AddCommand(OleMenuCommandService commands, int id, bool descending)
        {
            var command = new OleMenuCommand((sender, args) => Execute(descending), new CommandID(CommandSet, id));
            command.BeforeQueryStatus += (sender, args) =>
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                command.Enabled = TryGetSelection(out _, out _);
            };
            commands.AddCommand(command);
        }

        private static bool TryGetSelection(out IVsTextView view, out TextSpan span)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            view = null;
            span = new TextSpan();
            try
            {
                // The context menu owns focus. Query the last active native editor view,
                // not DTE.ActiveDocument.Selection during the menu's focus transition.
                var manager = Package.GetGlobalService(typeof(SVsTextManager)) as IVsTextManager;
                if (manager == null || ErrorHandler.Failed(manager.GetActiveView(0, null, out view)) || view == null)
                    return false;
                if (view.GetSelectionMode() != TextSelMode.SM_STREAM
                    || ErrorHandler.Failed(view.GetSelection(out int startLine, out int startColumn, out int endLine, out int endColumn)))
                    return false;
                if (startLine > endLine || (startLine == endLine && startColumn > endColumn))
                {
                    int line = startLine, column = startColumn;
                    startLine = endLine; startColumn = endColumn;
                    endLine = line; endColumn = column;
                }
                span = new TextSpan { iStartLine = startLine, iStartIndex = startColumn, iEndLine = endLine, iEndIndex = endColumn };
                return startLine != endLine || startColumn != endColumn;
            }
            catch (COMException) { return false; }
        }

        private void Execute(bool descending)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            bool vertical = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;
            try
            {
                if (!TryGetSelection(out var view, out var span)) return;
                ErrorHandler.ThrowOnFailure(view.GetBuffer(out var buffer));
                ErrorHandler.ThrowOnFailure(buffer.GetLineText(span.iStartLine, span.iStartIndex, span.iEndLine, span.iEndIndex, out string text));
                string sorted = SelectedTextSorter.Sort(text, descending, vertical);
                if (sorted == text) return;

                IntPtr replacement = Marshal.StringToHGlobalUni(sorted);
                try
                {
                    var changed = new TextSpan[1];
                    // A single buffer replacement keeps the operation undoable in one step.
                    ErrorHandler.ThrowOnFailure(buffer.ReplaceLines(span.iStartLine, span.iStartIndex, span.iEndLine, span.iEndIndex,
                        replacement, sorted.Length, changed));
                    view.SetSelection(changed[0].iStartLine, changed[0].iStartIndex, changed[0].iEndLine, changed[0].iEndIndex);
                }
                finally { Marshal.FreeHGlobal(replacement); }
            }
            catch (Exception ex)
            {
                VsShellUtilities.ShowMessageBox(package, ex.Message, "Sort Selected Text",
                    OLEMSGICON.OLEMSGICON_WARNING, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }
    }
}
