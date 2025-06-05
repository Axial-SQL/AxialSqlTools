using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;
using Microsoft.SqlServer.TransactSql.ScriptDom;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;


namespace AxialSqlTools
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class FormatQueryCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 4131;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("45457e02-6dec-4a4d-ab22-c9ee126d23c5");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="FormatQueryCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private FormatQueryCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static FormatQueryCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in FormatQueryCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new FormatQueryCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
                       
            ThreadHelper.ThrowIfNotOnUIThread();
            DTE dte = Package.GetGlobalService(typeof(DTE)) as DTE;

            if (dte?.ActiveDocument != null)
            {                

                try
                {
                    TextSelection selection = dte.ActiveDocument.Selection as TextSelection;

                    string existingCommandText = selection.Text.Trim();

                    if (!string.IsNullOrEmpty(existingCommandText))
                    {
                        string result = FormatCode(existingCommandText);
                        selection.Delete();
                        selection.Insert(result);
                        return;
                    }

                    // continue formatiing the entire document when nothing is selected                    
                    TextDocument textDoc = dte.ActiveDocument.Object("TextDocument") as TextDocument;
                    if (textDoc != null)
                    {
                        existingCommandText = textDoc.CreateEditPoint(textDoc.StartPoint).GetText(textDoc.EndPoint).Trim();

                        if (!string.IsNullOrEmpty(existingCommandText))
                        {
                            string result = FormatCode(existingCommandText);


                            EditPoint startPoint = textDoc.StartPoint.CreateEditPoint();
                            startPoint.ReplaceText(textDoc.EndPoint, result, (int)vsEPReplaceTextOptions.vsEPReplaceTextKeepMarkers);
                            
                            return;
                        }
                    }

                }
                catch (Exception ex)
                {

                    // Show a message box to prove we were here
                    VsShellUtilities.ShowMessageBox(
                        this.package,
                        ex.Message,
                        "Error parsing the code",
                        OLEMSGICON.OLEMSGICON_WARNING,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

                }
            }

        }


        class OwnVisitor : TSqlFragmentVisitor
        {
            public List<QualifiedJoin> QualifiedJoins = new List<QualifiedJoin>();
            public List<UnqualifiedJoin> UnqualifiedJoins = new List<UnqualifiedJoin>();
            public List<SearchedCaseExpression> CaseExpressions = new List<SearchedCaseExpression>();
            public List<StatementList> ProcBodies = new List<StatementList>();

            public override void ExplicitVisit(QualifiedJoin node)
            {                
                base.ExplicitVisit(node);

                QualifiedJoins.Add(node);

                // This is the source code from the Microsoft dll
                //GenerateFragmentIfNotNull(node.FirstTableReference);

                //GenerateNewLineOrSpace(_options.NewLineBeforeJoinClause);

                //GenerateQualifiedJoinType(node.QualifiedJoinType);

                //if (node.JoinHint != JoinHint.None)
                //{
                //    GenerateSpace();
                //    JoinHintHelper.Instance.GenerateSourceForOption(_writer, node.JoinHint);
                //}

                //GenerateSpaceAndKeyword(TSqlTokenType.Join);

                ////MarkClauseBodyAlignmentWhenNecessary(_options.NewlineBeforeJoinClause);

                //NewLine(); 
                //GenerateFragmentIfNotNull(node.SecondTableReference);

                //NewLine();
                //GenerateKeyword(TSqlTokenType.On);

                //GenerateSpaceAndFragmentIfNotNull(node.SearchCondition);

            }        

            public override void ExplicitVisit(UnqualifiedJoin node)
            {
                base.ExplicitVisit(node);

                UnqualifiedJoins.Add(node);

                // This is the source code from the Microsoft dll
                //GenerateFragmentIfNotNull(node.FirstTableReference);

                //List<TokenGenerator> generators = GetValueForEnumKey(_unqualifiedJoinTypeGenerators, node.UnqualifiedJoinType);
                //if (generators != null)
                //{
                //    GenerateSpace();
                //    GenerateTokenList(generators);
                //}

                //GenerateSpaceAndFragmentIfNotNull(node.SecondTableReference);
            }

            public override void ExplicitVisit(SearchedCaseExpression node)
            {
                base.ExplicitVisit(node);

                CaseExpressions.Add(node);

                // This is the source code from the Microsoft dll
                //GenerateKeyword(TSqlTokenType.Case);

                //GenerateSpaceAndFragmentIfNotNull(node.InputExpression);

                //foreach (SimpleWhenClause when in node.WhenClauses)
                //{
                //    GenerateSpaceAndFragmentIfNotNull(when);
                //}

                //if (node.ElseExpression != null)
                //{
                //    GenerateSpaceAndKeyword(TSqlTokenType.Else);
                //    GenerateSpaceAndFragmentIfNotNull(node.ElseExpression);
                //}

                //GenerateSpaceAndKeyword(TSqlTokenType.End);

                //GenerateSpaceAndCollation(node.Collation);
            }

            public override void ExplicitVisit(CreateProcedureStatement node)
            {
                base.ExplicitVisit(node);
                if (node.StatementList != null)
                {
                    ProcBodies.Add(node.StatementList);
                }
            }

            // NEW: collect each StatementList inside a CREATE FUNCTION
            public override void ExplicitVisit(CreateFunctionStatement node)
            {
                base.ExplicitVisit(node);
                if (node.StatementList != null)
                {
                    ProcBodies.Add(node.StatementList);
                }
            }

            // NEW: collect each StatementList inside a CREATE TRIGGER
            public override void ExplicitVisit(CreateTriggerStatement node)
            {
                base.ExplicitVisit(node);
                if (node.StatementList != null)
                {
                    ProcBodies.Add(node.StatementList);
                }
            }

            //
            // ─── NEW OVERRIDES FOR ALTER … ──────────────────────────────────────────────
            //
            public override void ExplicitVisit(AlterProcedureStatement node)
            {
                base.ExplicitVisit(node);
                if (node.StatementList != null)
                    ProcBodies.Add(node.StatementList);
            }

            public override void ExplicitVisit(AlterFunctionStatement node)
            {
                base.ExplicitVisit(node);
                if (node.StatementList != null)
                    ProcBodies.Add(node.StatementList);
            }

            public override void ExplicitVisit(AlterTriggerStatement node)
            {
                base.ExplicitVisit(node);
                if (node.StatementList != null)
                    ProcBodies.Add(node.StatementList);
            }

            //
            // ─── NEW OVERRIDES FOR CREATE OR ALTER … ────────────────────────────────────
            //
            public override void ExplicitVisit(CreateOrAlterProcedureStatement node)
            {
                base.ExplicitVisit(node);
                if (node.StatementList != null)
                    ProcBodies.Add(node.StatementList);
            }

            public override void ExplicitVisit(CreateOrAlterFunctionStatement node)
            {
                base.ExplicitVisit(node);
                if (node.StatementList != null)
                    ProcBodies.Add(node.StatementList);
            }

            public override void ExplicitVisit(CreateOrAlterTriggerStatement node)
            {
                base.ExplicitVisit(node);
                if (node.StatementList != null)
                    ProcBodies.Add(node.StatementList);
            }

        }

        /// <summary>
        /// Recursively walks a StatementList, forces a double‐newline between each sibling statement,
        /// then dives into any nested StatementList inside control‐flow blocks (BEGIN/END, IF, WHILE, etc.).
        /// </summary>
        /// <param name="stmtList">The StatementList to process.</param>
        /// <param name="sqlFragment">The root TSqlFragment, used to access the ScriptTokenStream.</param>
        private void InsertBlankLinesRecursive(StatementList stmtList, TSqlFragment sqlFragment)
        {
            // 1) Insert a blank line between each pair of sibling statements in this list
            var statements = stmtList.Statements;
            for (int i = 0; i < statements.Count - 1; i++)
            {              

                // Find the last token of the i‐th statement
                int endOfStmt = statements[i].LastTokenIndex;
                int nextIdx = endOfStmt + 1;

                if (nextIdx < sqlFragment.ScriptTokenStream.Count)
                {
                    TSqlParserToken nextToken = sqlFragment.ScriptTokenStream[nextIdx];
                    if (nextToken.TokenType == TSqlTokenType.WhiteSpace)
                    {
                        nextToken.Text = "\r\n\r\n";
                    }
                    else
                    {
                        // (In practice, ScriptGenerator always emits at least one whitespace token here,
                        // so you usually won’t hit this branch. If you do, you could insert a WhiteSpace
                        // token manually, but it’s rarely necessary.)
                    }
                }
            }

            // 2) For each statement in this list, check if it contains its own StatementList
            foreach (TSqlStatement stmt in statements)
            {
                // ─── Handle BEGIN ... END blocks ───────────────────────────────────────────
                // e.g.:
                //   CREATE PROCEDURE ... AS BEGIN
                //       <inner statements>
                //   END
                if (stmt is BeginEndBlockStatement beb && beb.StatementList != null)
                {
                    InsertBlankLinesRecursive(beb.StatementList, sqlFragment);
                }

                // ─── Handle IF ... THEN [ ... ] ELSE [ ... ] ──────────────────────────────
                // An IfStatement can have:
                //   • ThenStatement  (either a single statement or a StatementList or BeginEndBlock)
                //   • ElseStatement  (same possibilities)
                if (stmt is IfStatement ifStmt)
                {
                    //// THEN part
                    //if (ifStmt.ThenStatement is StatementList thenList)
                    //{
                    //    InsertBlankLinesRecursive(thenList, sqlFragment);
                    //}
                    //else
                    if (ifStmt.ThenStatement is BeginEndBlockStatement thenBlock && thenBlock.StatementList != null)
                    {
                        InsertBlankLinesRecursive(thenBlock.StatementList, sqlFragment);
                    }

                    //// ELSE part (if present)
                    //if (ifStmt.ElseStatement is StatementList elseList)
                    //{
                    //    InsertBlankLinesRecursive(elseList, sqlFragment);
                    //}
                    //else
                    if (ifStmt.ElseStatement is BeginEndBlockStatement elseBlock && elseBlock.StatementList != null)
                    {
                        InsertBlankLinesRecursive(elseBlock.StatementList, sqlFragment);
                    }
                }

                // ─── Handle WHILE loops ────────────────────────────────────────────────────
                //   WHILE <condition>
                //       <some statement or BEGIN...END>
                //if (stmt is WhileStatement whileStmt && whileStmt.Statement is StatementList whList)
                //{
                //    InsertBlankLinesRecursive(whList, sqlFragment);
                //}
                //else 
                if (stmt is WhileStatement ws && ws.Statement is BeginEndBlockStatement whBlock && whBlock.StatementList != null)
                {
                    InsertBlankLinesRecursive(whBlock.StatementList, sqlFragment);
                }

                // ─── Handle TRY...CATCH ───────────────────────────────────────────────────
                //   TRY
                //       <tryBlock>
                //   CATCH
                //       <catchBlock>
                if (stmt is TryCatchStatement tryCatch)
                {
                    // the TryBlock is either a StatementList or a BeginEndBlock
                    if (tryCatch.TryStatements is StatementList tryList)
                    {
                        InsertBlankLinesRecursive(tryList, sqlFragment);
                    }                    

                    if (tryCatch.CatchStatements is StatementList catchList)
                    {
                        InsertBlankLinesRecursive(catchList, sqlFragment);
                    }

                }

                //// ─── Handle FOR loops (BOTH FOR… AND FOR EACH CURSOR) ───────────────────────
                ////   FOR ... { <stmt> | BEGIN...END }
                ////   FOR EACH ... { <stmt> | BEGIN...END }
                //if (stmt is ForLoopStatementBase forStmt)
                //{
                //    if (forStmt.Statement is StatementList forList)
                //    {
                //        InsertBlankLinesRecursive(forList, sqlFragment);
                //    }
                //    else if (forStmt.Statement is BeginEndBlock forBlock && forBlock.StatementList != null)
                //    {
                //        InsertBlankLinesRecursive(forBlock.StatementList, sqlFragment);
                //    }
                //}

                //    // ─── (Optional) Handle other nested‐statement constructs ─────────────────────
                //    // If you also use:
                //    //   • USING (SQL Server 2022+)   → UsingStatement.StatementList
                //    //   • LOCK (SQL Server 2022+)    → LockStatement.StatementList
                //    //   • CURSOR declarations
                //    //   • MERGE … OUTPUT statements that embed nested blocks
                //    //
                //    // you can copy the same pattern: check “is UsingStatement” or “is LockStatement,”
                //    // grab its .StatementList or nested BeginEndBlock, and recurse.

                //    // Example for a USING statement (if you ever see it in your code):
                //    // if (stmt is UsingStatement usingStmt && usingStmt.Statement is StatementList usingList)
                //    // {
                //    //     InsertBlankLinesRecursive(usingList, sqlFragment);
                //    // }
                //    // else if (stmt is UsingStatement us && us.Statement is BeginEndBlock usingBlock)
                //    // {
                //    //     InsertBlankLinesRecursive(usingBlock.StatementList, sqlFragment);
                //    // }
            }
        }


        private string FormatCode(string oldCode)
        {
            string resultCode = "";

            TSql160Parser sqlParser = new TSql160Parser(false);

            IList<ParseError> parseErrors = new List<ParseError>();
            TSqlFragment result = sqlParser.Parse(new StringReader(oldCode), out parseErrors);

            if (parseErrors.Count > 0)
            {
                string errorStr = "";
                foreach (var strError in parseErrors)
                {
                    errorStr += Environment.NewLine + strError.Message;
                }

                throw new Exception($"TSqlParser unable to load selected T-SQL due to a syntax error:{Environment.NewLine}{errorStr}");
            }


            Sql160ScriptGenerator gen = new Sql160ScriptGenerator();
            gen.Options.AlignClauseBodies = false;
            //gen.Options.IncludeSemicolons = false;     
            gen.Options.SqlVersion = SqlVersion.Sql160; //TODO - try to get from current connection
            gen.GenerateScript(result, out resultCode);

            if (SettingsManager.GetApplyAdditionalCodeFormatting())
            {
                try
                {
                    resultCode = ApplySpecialFormat(resultCode, sqlParser);
                }
                catch
                {
                    //TODO - probably need to display the failure somehow
                }
            }

            return resultCode;           

        }

        private string ApplySpecialFormat(string oldCode, TSql160Parser sqlParser)
        {
            IList<ParseError> parseErrors = new List<ParseError>();

            TSqlFragment sqlFragment = sqlParser.Parse(new StringReader(oldCode), out parseErrors);

            OwnVisitor visitor = new OwnVisitor();
            sqlFragment.Accept(visitor);

            // visitor.TokenJoinLocations.Sort((a, b) => b.CompareTo(a));

            // special case #1 - remove new line after JOIN
            foreach (QualifiedJoin QJoin in visitor.QualifiedJoins)
            {

                int NextTokenNumber = QJoin.SecondTableReference.FirstTokenIndex;

                while (true)
                {

                    TSqlParserToken NextToken = sqlFragment.ScriptTokenStream[NextTokenNumber - 1];

                    if (NextToken.TokenType == TSqlTokenType.WhiteSpace)
                        if (NextToken.Text == "\r\n")
                            NextToken.Text = " ";
                        else if (NextToken.Text.Trim() == "")
                            NextToken.Text = "";

                    if (NextToken.TokenType == TSqlTokenType.Join)
                        break;

                    NextTokenNumber -= 1;

                    //just in case
                    if (NextTokenNumber < 0)
                        break;

                }

            }

            //special case #2 - JOIN .. ON -> add a tab before ON
            foreach (QualifiedJoin QJoin in visitor.QualifiedJoins)
            {

                int NextTokenNumber = QJoin.SearchCondition.FirstTokenIndex;

                while (true)
                {

                    TSqlParserToken NextToken = sqlFragment.ScriptTokenStream[NextTokenNumber];

                    if (NextToken.TokenType == TSqlTokenType.On)
                    { // replace previos white-space with the new line and a number of spaces for offset

                        TSqlParserToken PreviousToken = sqlFragment.ScriptTokenStream[NextTokenNumber - 1];
                        if (PreviousToken.TokenType == TSqlTokenType.WhiteSpace)
                        {
                            PreviousToken.Text = PreviousToken.Text + new string(' ', 4);
                            break;
                        }

                    }

                    NextTokenNumber -= 1;

                    //just in case
                    if (NextTokenNumber < 0)
                        break;

                }
            }

            //special case #3 - CROSS should be on the new line
            foreach (UnqualifiedJoin CrossJoin in visitor.UnqualifiedJoins)
            {
                int NextTokenNumber = CrossJoin.SecondTableReference.FirstTokenIndex;

                while (true)
                {

                    TSqlParserToken NextToken = sqlFragment.ScriptTokenStream[NextTokenNumber];

                    if (NextToken.TokenType == TSqlTokenType.Cross)
                    { // replace previos white-space with the new line and a number of spaces for offset

                        TSqlParserToken PreviousToken = sqlFragment.ScriptTokenStream[NextTokenNumber - 1];
                        if (PreviousToken.TokenType == TSqlTokenType.WhiteSpace)
                        {
                            PreviousToken.Text = "\r\n" + new string(' ', CrossJoin.StartColumn - 1);
                            break;
                        }

                    }

                    NextTokenNumber -= 1;

                    //just in case
                    if (NextTokenNumber < 0)
                        break;

                }

            }

            //special case #4 - CASE <new line + tab> WHEN <new line + tab + tab> THEN <new line + tab> ELSE <new line> END
            foreach (SearchedCaseExpression CaseExpr in visitor.CaseExpressions)
            {
                // add new line and spaces+4 before WHEN
                foreach (WhenClause WC in CaseExpr.WhenClauses)
                {
                    int FirstTokenNumber = WC.FirstTokenIndex;

                    int WhenIdent = 0;

                    while (true)
                    {
                        TSqlParserToken NextToken = sqlFragment.ScriptTokenStream[FirstTokenNumber];

                        if (NextToken.TokenType == TSqlTokenType.WhiteSpace)
                        { // replace previos white-space with the new line and a number of spaces for offset
                            NextToken.Text = "\r\n" + new string(' ', CaseExpr.StartColumn + 4);

                            WhenIdent = CaseExpr.StartColumn + 4;

                            break;
                        }

                        FirstTokenNumber -= 1;

                        //just in case
                        if (FirstTokenNumber < 0)
                            break;
                    }

                    //multi-line expression inside WHEN might be too far to the right, move it to the left
                    if (WhenIdent > 0 && WhenIdent != WC.StartColumn)
                    {

                        var FirshWhenToken = WC.FirstTokenIndex;

                        while (FirshWhenToken < WC.LastTokenIndex)
                        {
                            TSqlParserToken NextThenToken = sqlFragment.ScriptTokenStream[FirshWhenToken];

                            if (NextThenToken.TokenType == TSqlTokenType.WhiteSpace && NextThenToken.Column == 1)
                            {
                                NextThenToken.Text = new string(' ', WhenIdent + 5);
                            }

                            FirshWhenToken += 1;
                        }

                    }


                }

                // add new line and spaces+8 before THEN
                foreach (WhenClause WC in CaseExpr.WhenClauses)
                {
                    int FirstTokenNumber = WC.ThenExpression.FirstTokenIndex - 3;

                    int ThenIdent = 0;

                    while (true)
                    {
                        TSqlParserToken NextToken = sqlFragment.ScriptTokenStream[FirstTokenNumber];

                        if (NextToken.TokenType == TSqlTokenType.WhiteSpace)
                        { // replace previos white-space with the new line and a number of spaces for offset
                            NextToken.Text = "\r\n" + new string(' ', CaseExpr.StartColumn + 8);

                            ThenIdent = CaseExpr.StartColumn + 8;

                            break;
                        }

                        FirstTokenNumber -= 1;

                        //just in case
                        if (FirstTokenNumber < 0)
                            break;
                    }

                    //multi-line expression inside THEN might be too far to the right, move it to the left
                    if (ThenIdent > 0 && ThenIdent != WC.ThenExpression.StartColumn)
                    {

                        var FirshThenToken = WC.ThenExpression.FirstTokenIndex;

                        while (FirshThenToken < WC.ThenExpression.LastTokenIndex)
                        {
                            TSqlParserToken NextThenToken = sqlFragment.ScriptTokenStream[FirshThenToken];

                            if (NextThenToken.TokenType == TSqlTokenType.WhiteSpace && NextThenToken.Column == 1)
                            {
                                NextThenToken.Text = new string(' ', ThenIdent + 5);
                            }

                            FirshThenToken += 1;
                        }

                    }
                }

                // add new line and spaces+4 before ELSE
                if (CaseExpr.ElseExpression != null)
                {
                    int FirstTokenNumber = CaseExpr.ElseExpression.FirstTokenIndex - 3;

                    int ElseIdent = 0;

                    while (true)
                    {
                        TSqlParserToken NextToken = sqlFragment.ScriptTokenStream[FirstTokenNumber];

                        if (NextToken.TokenType == TSqlTokenType.WhiteSpace)
                        { // replace previos white-space with the new line and a number of spaces for offset
                            NextToken.Text = "\r\n" + new string(' ', CaseExpr.StartColumn + 4);
                            ElseIdent = CaseExpr.StartColumn + 4;
                            break;
                        }

                        FirstTokenNumber -= 1;

                        //just in case
                        if (FirstTokenNumber < 0)
                            break;
                    }

                    //multi-line expression inside ELSE might be too far to the right, move it to the left
                    if (ElseIdent > 0 && ElseIdent != CaseExpr.ElseExpression.StartColumn)
                    {

                        var FirshWhenToken = CaseExpr.ElseExpression.FirstTokenIndex;

                        while (FirshWhenToken < CaseExpr.ElseExpression.LastTokenIndex)
                        {
                            TSqlParserToken NextThenToken = sqlFragment.ScriptTokenStream[FirshWhenToken];

                            if (NextThenToken.TokenType == TSqlTokenType.WhiteSpace && NextThenToken.Column == 1)
                            {
                                NextThenToken.Text = new string(' ', ElseIdent + 5);
                            }

                            FirshWhenToken += 1;
                        }

                    }

                }

                // add new line and spaces before END
                int LastTokenNumber = CaseExpr.LastTokenIndex;

                while (true)
                {

                    TSqlParserToken NextToken = sqlFragment.ScriptTokenStream[LastTokenNumber];

                    if (NextToken.TokenType == TSqlTokenType.WhiteSpace)
                    { // replace previos white-space with the new line and a number of spaces for offset
                        NextToken.Text = "\r\n" + new string(' ', CaseExpr.StartColumn - 1);
                        break;
                    }

                    LastTokenNumber -= 1;

                    //just in case
                    if (LastTokenNumber < 0)
                        break;
                }

            }

            //special case #5 - add two new lines after each PROC/FUNC/TRIGGER statement
            foreach (StatementList topLevel in visitor.ProcBodies)
            {
                InsertBlankLinesRecursive(topLevel, sqlFragment);
            }

            // return full recompiled result
            StringBuilder sqlText = new StringBuilder();
            foreach (var Token in sqlFragment.ScriptTokenStream)
            {
                sqlText.Append(Token.Text);
            }

            return sqlText.ToString();
        }

    }
}
