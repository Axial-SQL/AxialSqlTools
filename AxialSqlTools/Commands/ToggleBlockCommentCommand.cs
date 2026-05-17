using System;
using System.ComponentModel.Design;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace AxialSqlTools
{
    internal sealed class ToggleBlockCommentCommand
    {
        public const int CommandId = 4155;

        public static readonly Guid CommandSet = new Guid("45457e02-6dec-4a4d-ab22-c9ee126d23c5");

        private readonly AsyncPackage package;

        private ToggleBlockCommentCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static ToggleBlockCommentCommand Instance { get; private set; }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new ToggleBlockCommentCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            DTE dte = Package.GetGlobalService(typeof(DTE)) as DTE;
            if (dte?.ActiveDocument == null)
                return;

            try
            {
                TextDocument textDoc = dte.ActiveDocument.Object("TextDocument") as TextDocument;
                TextSelection selection = dte.ActiveDocument.Selection as TextSelection;
                if (textDoc == null || selection == null)
                    return;

                if (!selection.IsEmpty)
                {
                    ToggleSelectedText(selection);
                    return;
                }

                ToggleCommentAtCursor(textDoc, selection);
            }
            catch (Exception ex)
            {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    ex.Message,
                    "Error toggling block comment",
                    OLEMSGICON.OLEMSGICON_WARNING,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        private void ToggleSelectedText(TextSelection selection)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string selectedText = selection.Text;
            if (string.IsNullOrEmpty(selectedText))
                return;

            if (TryGetOuterBlockCommentDelimiters(selectedText, out int openIndex, out int closeIndex))
            {
                string uncommentedText = selectedText.Remove(closeIndex, 2).Remove(openIndex, 2);
                selection.Delete();
                selection.Insert(uncommentedText, (int)vsInsertFlags.vsInsertFlagsContainNewText);
                return;
            }

            selection.Delete();
            selection.Insert("/*" + selectedText + "*/", (int)vsInsertFlags.vsInsertFlagsContainNewText);
        }

        private void ToggleCommentAtCursor(TextDocument textDoc, TextSelection selection)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string fullText = textDoc.CreateEditPoint(textDoc.StartPoint).GetText(textDoc.EndPoint);
            if (string.IsNullOrEmpty(fullText))
                return;

            int originalOffset = selection.ActivePoint.AbsoluteCharOffset;
            int caretIndex = Math.Max(0, Math.Min(fullText.Length, originalOffset - 1));
            if (!TryFindEnclosingBlockComment(fullText, caretIndex, out int openIndex, out int closeIndex))
                return;

            EditPoint closePoint = textDoc.StartPoint.CreateEditPoint();
            closePoint.MoveToAbsoluteOffset(closeIndex + 1);
            closePoint.Delete(2);

            EditPoint openPoint = textDoc.StartPoint.CreateEditPoint();
            openPoint.MoveToAbsoluteOffset(openIndex + 1);
            openPoint.Delete(2);

            int restoredOffset = originalOffset;
            if (caretIndex > closeIndex)
            {
                restoredOffset -= 4;
            }
            else if (caretIndex > openIndex)
            {
                restoredOffset -= 2;
            }

            selection.MoveToAbsoluteOffset(Math.Max(1, restoredOffset), false);
        }

        private bool TryGetOuterBlockCommentDelimiters(string text, out int openIndex, out int closeIndex)
        {
            openIndex = -1;
            closeIndex = -1;

            int firstNonWhitespace = 0;
            while (firstNonWhitespace < text.Length && char.IsWhiteSpace(text[firstNonWhitespace]))
            {
                firstNonWhitespace++;
            }

            int lastNonWhitespace = text.Length - 1;
            while (lastNonWhitespace >= firstNonWhitespace && char.IsWhiteSpace(text[lastNonWhitespace]))
            {
                lastNonWhitespace--;
            }

            if (lastNonWhitespace - firstNonWhitespace + 1 < 4)
                return false;

            if (text[firstNonWhitespace] != '/' || text[firstNonWhitespace + 1] != '*')
                return false;

            if (text[lastNonWhitespace - 1] != '*' || text[lastNonWhitespace] != '/')
                return false;

            openIndex = firstNonWhitespace;
            closeIndex = lastNonWhitespace - 1;
            return true;
        }

        private bool TryFindEnclosingBlockComment(string text, int caretIndex, out int openIndex, out int closeIndex)
        {
            openIndex = -1;
            closeIndex = -1;

            int searchStart = Math.Min(caretIndex, text.Length - 1);
            openIndex = text.LastIndexOf("/*", searchStart, StringComparison.Ordinal);
            if (openIndex < 0)
                return false;

            int lastCloseBeforeCaret = text.LastIndexOf("*/", searchStart, StringComparison.Ordinal);
            if (lastCloseBeforeCaret > openIndex)
                return false;

            closeIndex = text.IndexOf("*/", openIndex + 2, StringComparison.Ordinal);
            if (closeIndex < 0 || caretIndex > closeIndex + 2)
                return false;

            return true;
        }
    }
}
