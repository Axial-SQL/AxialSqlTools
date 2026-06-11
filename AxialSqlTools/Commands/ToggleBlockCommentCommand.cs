using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using EnvDTE;
using Microsoft.SqlServer.TransactSql.ScriptDom;
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

            TextPoint activePoint = selection.ActivePoint;
            var caret = new TextPosition(activePoint.Line, activePoint.LineCharOffset);
            if (!TryFindEnclosingBlockComment(fullText, caret, out TextPosition openPosition, out TextPosition closePosition))
                return;

            EditPoint closePoint = textDoc.StartPoint.CreateEditPoint();
            closePoint.MoveToLineAndOffset(closePosition.Line, closePosition.Column);
            closePoint.Delete(2);

            EditPoint openPoint = textDoc.StartPoint.CreateEditPoint();
            openPoint.MoveToLineAndOffset(openPosition.Line, openPosition.Column);
            openPoint.Delete(2);
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

        private bool TryFindEnclosingBlockComment(
            string text,
            TextPosition caret,
            out TextPosition openPosition,
            out TextPosition closePosition)
        {
            openPosition = default(TextPosition);
            closePosition = default(TextPosition);

            var parser = new TSql170Parser(initialQuotedIdentifiers: true);
            IList<ParseError> errors;
            IList<TSqlParserToken> tokens = parser.GetTokenStream(new StringReader(text), out errors);
            foreach (TSqlParserToken token in tokens)
            {
                if (token.TokenType != TSqlTokenType.MultilineComment || string.IsNullOrEmpty(token.Text))
                    continue;

                var tokenStart = new TextPosition(token.Line, token.Column);
                TextPosition tokenEnd = GetPositionAfterText(token.Text, tokenStart);
                if (!ContainsPosition(tokenStart, tokenEnd, caret))
                    continue;

                if (!TryGetTextIndexAtPosition(token.Text, tokenStart, caret, out int caretIndex))
                    continue;

                if (!TryFindEnclosingBlockCommentDelimiters(token.Text, caretIndex, out int openIndex, out int closeIndex))
                    continue;

                openPosition = GetPositionAtTextIndex(token.Text, tokenStart, openIndex);
                closePosition = GetPositionAtTextIndex(token.Text, tokenStart, closeIndex);
                return true;
            }

            return false;
        }

        private bool TryFindEnclosingBlockCommentDelimiters(string commentText, int caretIndex, out int openIndex, out int closeIndex)
        {
            openIndex = -1;
            closeIndex = -1;

            var blockStarts = new Stack<int>();
            int bestOpenIndex = -1;
            int bestCloseIndex = -1;

            for (int i = 0; i < commentText.Length - 1; i++)
            {
                if (commentText[i] == '/' && commentText[i + 1] == '*')
                {
                    blockStarts.Push(i);
                    i++;
                    continue;
                }

                if (commentText[i] == '*' && commentText[i + 1] == '/' && blockStarts.Count > 0)
                {
                    int candidateOpenIndex = blockStarts.Pop();
                    int candidateCloseIndex = i;
                    if (candidateOpenIndex <= caretIndex
                        && caretIndex < candidateCloseIndex + 2
                        && candidateOpenIndex > bestOpenIndex)
                    {
                        bestOpenIndex = candidateOpenIndex;
                        bestCloseIndex = candidateCloseIndex;
                    }

                    i++;
                }
            }

            if (bestOpenIndex < 0)
                return false;

            openIndex = bestOpenIndex;
            closeIndex = bestCloseIndex;
            return true;
        }

        private bool ContainsPosition(TextPosition start, TextPosition end, TextPosition position)
        {
            return ComparePositions(start, position) <= 0 && ComparePositions(position, end) < 0;
        }

        private int ComparePositions(TextPosition left, TextPosition right)
        {
            if (left.Line != right.Line)
                return left.Line.CompareTo(right.Line);

            return left.Column.CompareTo(right.Column);
        }

        private bool TryGetTextIndexAtPosition(string text, TextPosition start, TextPosition position, out int index)
        {
            TextPosition current = start;
            for (int i = 0; i <= text.Length; i++)
            {
                if (current.Line == position.Line && current.Column == position.Column)
                {
                    index = i;
                    return true;
                }

                if (i == text.Length)
                    break;

                AdvancePosition(text, ref i, ref current);
            }

            index = -1;
            return false;
        }

        private TextPosition GetPositionAtTextIndex(string text, TextPosition start, int targetIndex)
        {
            TextPosition current = start;
            for (int i = 0; i < targetIndex && i < text.Length; i++)
            {
                AdvancePosition(text, ref i, ref current);
            }

            return current;
        }

        private TextPosition GetPositionAfterText(string text, TextPosition start)
        {
            return GetPositionAtTextIndex(text, start, text.Length);
        }

        private void AdvancePosition(string text, ref int index, ref TextPosition position)
        {
            if (text[index] == '\r')
            {
                if (index + 1 < text.Length && text[index + 1] == '\n')
                {
                    index++;
                }

                position = new TextPosition(position.Line + 1, 1);
                return;
            }

            if (text[index] == '\n')
            {
                position = new TextPosition(position.Line + 1, 1);
                return;
            }

            position = new TextPosition(position.Line, position.Column + 1);
        }

        private struct TextPosition
        {
            public TextPosition(int line, int column)
            {
                Line = line;
                Column = column;
            }

            public int Line { get; }

            public int Column { get; }
        }
    }
}
