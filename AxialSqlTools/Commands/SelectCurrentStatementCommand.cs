using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.SqlServer.TransactSql.ScriptDom;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace AxialSqlTools
{
    internal sealed class SelectCurrentStatementCommand
    {
        public const int CommandId = 4152;

        public static readonly Guid CommandSet = new Guid("45457e02-6dec-4a4d-ab22-c9ee126d23c5");

        private readonly AsyncPackage package;

        private SelectCurrentStatementCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static SelectCurrentStatementCommand Instance { get; private set; }

        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get { return this.package; }
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new SelectCurrentStatementCommand(package, commandService);
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
                if (textDoc == null)
                    return;

                TextSelection selection = dte.ActiveDocument.Selection as TextSelection;
                if (selection == null)
                    return;

                int cursorLine = selection.ActivePoint.Line;
                int cursorColumn = selection.ActivePoint.DisplayColumn;

                string fullText = textDoc.CreateEditPoint(textDoc.StartPoint).GetText(textDoc.EndPoint);

                if (string.IsNullOrEmpty(fullText))
                    return;

                // Try parsing with TSql170Parser
                if (TrySelectWithParser(fullText, cursorLine, cursorColumn, selection))
                    return;

                // Fallback: GO-based batch detection
                TrySelectWithFallback(fullText, cursorLine, cursorColumn, selection);
            }
            catch (Exception ex)
            {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    ex.Message,
                    "Error selecting current statement",
                    OLEMSGICON.OLEMSGICON_WARNING,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        private bool TrySelectWithParser(string fullText, int cursorLine, int cursorColumn, TextSelection selection)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            TSql170Parser sqlParser = new TSql170Parser(false);
            IList<ParseError> parseErrors;
            TSqlFragment fragment = sqlParser.Parse(new StringReader(fullText), out parseErrors);

            // Don't bail on parse errors — the parser still returns a usable AST
            // with successfully parsed statements even when some have syntax errors
            TSqlScript script = fragment as TSqlScript;
            if (script == null)
                return false;

            // Track the best (top-level) statement match
            int bestStartLine = -1;
            int bestStartCol = -1;
            int bestEndLine = -1;
            int bestEndCol = -1;

            foreach (TSqlBatch batch in script.Batches)
            {
                foreach (TSqlStatement stmt in batch.Statements)
                {
                    var firstToken = fragment.ScriptTokenStream[stmt.FirstTokenIndex];
                    int startLine = firstToken.Line;
                    int startCol = firstToken.Column; // ScriptDom columns are already 1-based

                    var lastToken = fragment.ScriptTokenStream[stmt.LastTokenIndex];
                    int endLine = lastToken.Line;
                    int endCol = lastToken.Column + lastToken.Text.Length; // already 1-based, points past last char

                    if (CursorIsWithin(cursorLine, cursorColumn, startLine, startCol, endLine, endCol))
                    {
                        bestStartLine = startLine;
                        bestStartCol = startCol;
                        bestEndLine = endLine;
                        bestEndCol = endCol;
                    }
                }
            }

            if (bestStartLine < 0)
                return true; // Parser worked but cursor is not inside any statement — do nothing

            // Select the statement in the editor
            selection.MoveToLineAndOffset(bestStartLine, bestStartCol, false);
            selection.MoveToLineAndOffset(bestEndLine, bestEndCol, true);
            return true;
        }

        private bool CursorIsWithin(int cursorLine, int cursorCol, int startLine, int startCol, int endLine, int endCol)
        {
            // Check if cursor is before the start
            if (cursorLine < startLine || (cursorLine == startLine && cursorCol < startCol))
                return false;

            // Check if cursor is after the end
            if (cursorLine > endLine || (cursorLine == endLine && cursorCol > endCol))
                return false;

            return true;
        }

        private void TrySelectWithFallback(string fullText, int cursorLine, int cursorColumn, TextSelection selection)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Split text into lines for processing
            string[] lines = fullText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            // Find batch boundaries using GO
            var goPattern = new Regex(@"^\s*GO\s*$", RegexOptions.IgnoreCase);

            int batchStartLine = 1; // 1-based
            int batchEndLine = lines.Length;
            bool foundBatch = false;

            // Find which batch the cursor is in
            int currentLine = 1;
            int currentBatchStart = 1;

            for (int i = 0; i < lines.Length; i++)
            {
                currentLine = i + 1; // 1-based

                if (goPattern.IsMatch(lines[i]))
                {
                    if (cursorLine >= currentBatchStart && cursorLine < currentLine)
                    {
                        batchStartLine = currentBatchStart;
                        batchEndLine = currentLine - 1;
                        foundBatch = true;
                        break;
                    }
                    currentBatchStart = currentLine + 1;
                }
            }

            if (!foundBatch)
            {
                // Cursor is in the last batch (after last GO or no GO at all)
                if (cursorLine >= currentBatchStart)
                {
                    batchStartLine = currentBatchStart;
                    batchEndLine = lines.Length;
                }
                else
                {
                    return; // cursor on a GO line itself
                }
            }

            // Within the batch, find the statement segment using semicolons
            // Simple approach: split by semicolons at line level
            int stmtStartLine = batchStartLine;
            int stmtEndLine = batchEndLine;

            for (int i = batchStartLine - 1; i < batchEndLine; i++)
            {
                string line = lines[i];
                int lineNum = i + 1;

                // Check if line contains a semicolon (simple heuristic, doesn't account for strings)
                if (line.Contains(";"))
                {
                    if (lineNum < cursorLine)
                    {
                        stmtStartLine = lineNum + 1;
                    }
                    else if (lineNum == cursorLine)
                    {
                        // Cursor is on a line with semicolon - include this line
                        stmtEndLine = lineNum;
                        break;
                    }
                    else
                    {
                        stmtEndLine = lineNum;
                        break;
                    }
                }
            }

            // Trim leading/trailing empty lines from the selection
            while (stmtStartLine <= stmtEndLine && string.IsNullOrWhiteSpace(lines[stmtStartLine - 1]))
                stmtStartLine++;

            while (stmtEndLine >= stmtStartLine && string.IsNullOrWhiteSpace(lines[stmtEndLine - 1]))
                stmtEndLine--;

            if (stmtStartLine > stmtEndLine)
                return; // nothing to select

            // Select in editor
            int startOffset = 1;
            int endLineLength = lines[stmtEndLine - 1].Length + 1; // +1 for 1-based offset past last char

            selection.MoveToLineAndOffset(stmtStartLine, startOffset, false);
            selection.MoveToLineAndOffset(stmtEndLine, endLineLength, true);
        }
    }
}
