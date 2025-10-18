using Microsoft.SqlServer.TransactSql.ScriptDom;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using static AxialSqlTools.AxialSqlToolsPackage;

namespace AxialSqlTools
{
    public static class TSqlFormatter
    {
        class OwnVisitor : TSqlFragmentVisitor
        {
            public List<QualifiedJoin> QualifiedJoins = new List<QualifiedJoin>();
            public List<UnqualifiedJoin> UnqualifiedJoins = new List<UnqualifiedJoin>();
            public List<SearchedCaseExpression> CaseExpressions = new List<SearchedCaseExpression>();
            public List<StatementList> ProcBodies = new List<StatementList>();
            public List<ExecuteStatement> ExecStatements = new List<ExecuteStatement>();
            public List<FunctionCall> FunctionCalls = new List<FunctionCall>();
            public List<BeginEndBlockStatement> BeginEndBlocks = new List<BeginEndBlockStatement>();
            public List<DeclareVariableStatement> DeclareStatements = new List<DeclareVariableStatement>();
            public List<CreateProcedureStatement> SprocDefinitionsCreate = new List<CreateProcedureStatement>();
            public List<AlterProcedureStatement> SprocDefinitionsAlter = new List<AlterProcedureStatement>();
            public List<CreateOrAlterProcedureStatement> SprocDefinitionsCreateAlter = new List<CreateOrAlterProcedureStatement>();
            public List<QuerySpecification> SelectsWithTop = new List<QuerySpecification>();

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
                SprocDefinitionsCreate.Add(node);
                if (node.StatementList != null)
                    ProcBodies.Add(node.StatementList);
            }

            public override void ExplicitVisit(CreateFunctionStatement node)
            {
                base.ExplicitVisit(node);
                if (node.StatementList != null)
                    ProcBodies.Add(node.StatementList);
            }

            public override void ExplicitVisit(CreateTriggerStatement node)
            {
                base.ExplicitVisit(node);
                if (node.StatementList != null)
                    ProcBodies.Add(node.StatementList);
            }

            public override void ExplicitVisit(AlterProcedureStatement node)
            {
                base.ExplicitVisit(node);
                SprocDefinitionsAlter.Add(node);
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

            public override void ExplicitVisit(CreateOrAlterProcedureStatement node)
            {
                base.ExplicitVisit(node);
                SprocDefinitionsCreateAlter.Add(node);
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

            public override void ExplicitVisit(ExecuteStatement node)
            {
                base.ExplicitVisit(node);
                ExecStatements.Add(node);
            }

            public override void ExplicitVisit(FunctionCall node)
            {
                base.ExplicitVisit(node);
                FunctionCalls.Add(node);
            }

            public override void ExplicitVisit(BeginEndBlockStatement node)
            {
                base.ExplicitVisit(node);
                BeginEndBlocks.Add(node);
            }

            public override void ExplicitVisit(DeclareVariableStatement node)
            {
                base.ExplicitVisit(node);
                DeclareStatements.Add(node);
            }

            public override void ExplicitVisit(QuerySpecification node)
            {
                base.ExplicitVisit(node);
                if (node.TopRowFilter != null)
                    SelectsWithTop.Add(node);
            }
        }

        /// <summary>
        /// Recursively walks a StatementList, forces a double‐newline between each sibling statement,
        /// then dives into any nested StatementList inside control‐flow blocks (BEGIN/END, IF, WHILE, etc.).
        /// </summary>
        /// <param name="stmtList">The StatementList to process.</param>
        /// <param name="sqlFragment">The root TSqlFragment, used to access the ScriptTokenStream.</param>
        private static void InsertBlankLinesRecursive(StatementList stmtList, TSqlFragment sqlFragment, bool skipTopLevel = false)
        {

            // 1) Insert a blank line between each pair of sibling statements in this list
            var statements = stmtList.Statements;
            if (!skipTopLevel)
            {
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
            }

            // 2) For each statement in this list, check if it contains its own StatementList
            foreach (TSqlStatement stmt in statements)
            {
                // ─── Handle BEGIN ... END blocks ───────────────────────────────────────────
                if (stmt is BeginEndBlockStatement beb && beb.StatementList != null)
                {
                    InsertBlankLinesRecursive(beb.StatementList, sqlFragment);
                }

                // ─── Handle IF ... THEN [ ... ] ELSE [ ... ] ──────────────────────────────
                if (stmt is IfStatement ifStmt)
                {
                    if (ifStmt.ThenStatement is BeginEndBlockStatement thenBlock && thenBlock.StatementList != null)
                    {
                        InsertBlankLinesRecursive(thenBlock.StatementList, sqlFragment);
                    }

                    if (ifStmt.ElseStatement is BeginEndBlockStatement elseBlock && elseBlock.StatementList != null)
                    {
                        InsertBlankLinesRecursive(elseBlock.StatementList, sqlFragment);
                    }
                }

                // ─── Handle WHILE loops ────────────────────────────────────────────────────
                if (stmt is WhileStatement ws && ws.Statement is BeginEndBlockStatement whBlock && whBlock.StatementList != null)
                {
                    InsertBlankLinesRecursive(whBlock.StatementList, sqlFragment);
                }

                // ─── Handle TRY...CATCH ───────────────────────────────────────────────────
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

            }
        }


        public static string FormatCode(string oldCode, SettingsManager.TSqlCodeFormatSettings settingsOverride = null)
        {
            string resultCode = "";

            TSql170Parser sqlParser = new TSql170Parser(false);

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

            var formatSettings = SettingsManager.GetTSqlCodeFormatSettings();
            if (settingsOverride != null)
            {
                formatSettings = settingsOverride;
            }

            Sql170ScriptGenerator gen = new Sql170ScriptGenerator();
            gen.Options.AlignClauseBodies = false;
            gen.Options.SqlVersion = SqlVersion.Sql170; //TODO - try to get from current connection

            if (formatSettings.preserveComments)
            {
                resultCode = TsqlFormatterCommentInterleaver.GenerateWithComments(result, gen, sqlParser);
            }
            else 
            { 
                gen.GenerateScript(result, out resultCode);
            }

            if (formatSettings.HasAnyFormattingEnabled())
            {
                try
                {
                    resultCode = ApplySpecialFormat(resultCode, sqlParser, formatSettings);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "An error occurred while applying special formatting to the code.");
                }
            }

            return resultCode;

        }

        private static string ApplySpecialFormat(string oldCode, TSql170Parser sqlParser, SettingsManager.TSqlCodeFormatSettings formatSettings)
        {
            IList<ParseError> parseErrors = new List<ParseError>();

            TSqlFragment sqlFragment = sqlParser.Parse(new StringReader(oldCode), out parseErrors);

            OwnVisitor visitor = new OwnVisitor();
            sqlFragment.Accept(visitor);

            // special case #1 - remove new line after JOIN
            if (formatSettings.removeNewLineAfterJoin)
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
            if (formatSettings.addTabAfterJoinOn)
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
            if (formatSettings.moveCrossJoinToNewLine)
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
            if (formatSettings.formatCaseAsMultiline)
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

            //special case #5.1 - add two new lines after each PROC/FUNC/TRIGGER statement
            //special case #5.2 - also add blank lines in every anonymous batch
            if (formatSettings.addNewLineBetweenStatementsInBlocks)
            {                 
                foreach (StatementList topLevel in visitor.ProcBodies)
                {
                    InsertBlankLinesRecursive(topLevel, sqlFragment);
                }
                
                if (sqlFragment is TSqlScript script)
                {
                    foreach (var batch in script.Batches)
                    {
                        // build a StatementList wrapper around the batch's statements
                        var anonList = new StatementList { Statements = { } };
                        foreach (var stmt in batch.Statements)
                            anonList.Statements.Add(stmt);

                        InsertBlankLinesRecursive(anonList, sqlFragment, skipTopLevel: true);
                    }
                }
            }

            // special case #6 – break sproc parameters onto separate lines
            if (formatSettings.breakSprocParametersPerLine)
            {
                var tokens = sqlFragment.ScriptTokenStream;

                // for every EXEC … call
                foreach (var execStmt in visitor.ExecStatements)
                {
                    var spec = execStmt.ExecuteSpecification;
                    var execEntry = (ExecutableProcedureReference)spec.ExecutableEntity;

                    if (execEntry.Parameters.Count > 1)
                    {
                        // 0) figure out how much indent EXEC itself already has
                        string baseIndent = "";
                        int wsIdxBeforeExec = execStmt.FirstTokenIndex - 1;
                        if (wsIdxBeforeExec >= 0 && tokens[wsIdxBeforeExec].TokenType == TSqlTokenType.WhiteSpace)
                        {
                            var wsText = tokens[wsIdxBeforeExec].Text;
                            // grab whatever is after the last newline
                            int lastNl = wsText.LastIndexOf("\r\n");
                            baseIndent = lastNl >= 0
                                ? wsText.Substring(lastNl + 2)
                                : wsText;
                        }

                        // 1) put first param on its own indented line
                        int insertPos = execEntry.ProcedureReference.LastTokenIndex + 1;
                        if (insertPos < tokens.Count && tokens[insertPos].TokenType == TSqlTokenType.WhiteSpace)
                            tokens[insertPos].Text = "\r\n"
                                                    + baseIndent
                                                    + "\t";

                        // 2) for every comma between params, break+indent by the same amount
                        for (int i = spec.FirstTokenIndex; i <= spec.LastTokenIndex; i++)
                        {
                            if (tokens[i].TokenType == TSqlTokenType.Comma)
                            {
                                int wsAfterComma = i + 1;
                                if (wsAfterComma < tokens.Count
                                 && tokens[wsAfterComma].TokenType == TSqlTokenType.WhiteSpace)
                                {
                                    tokens[wsAfterComma].Text = " "         // space after comma
                                                             + "\r\n"
                                                             + baseIndent
                                                             + "\t";
                                }
                            }
                        }
                    }
                }
            }

            // special case #7 - uppercase built-in function names
            if (formatSettings.uppercaseBuiltInFunctions)
            {
                var tokens = sqlFragment.ScriptTokenStream;
                foreach (var func in visitor.FunctionCalls)
                {
                    // skip if schema-qualified (e.g. dbo.MyFunc)
                    if (func.CallTarget != null)
                        continue;

                    // FunctionName spans one or more identifier tokens immediately before "("
                    int start = func.FunctionName.FirstTokenIndex;
                    int end = func.FunctionName.LastTokenIndex;

                    for (int i = start; i <= end; i++)
                    {
                        var tok = tokens[i];
                        if (tok.TokenType == TSqlTokenType.Identifier)
                            tok.Text = tok.Text.ToUpperInvariant();
                    }
                }
            }

            // special case #8 – do not indent BEGIN/END in control‐flow
            if (formatSettings.unindentBeginEndBlocks)
            {
                // adjust to "\t" if you indent with tabs
                string indentString = new string(' ', 4);

                foreach (var beb in visitor.BeginEndBlocks)
                {
                    // include the whitespace before BEGIN (at FirstTokenIndex-1)
                    // and every whitespace token up through the token before END
                    int startIdx = Math.Max(0, beb.FirstTokenIndex - 1);
                    int endIdx = Math.Max(0, beb.LastTokenIndex - 1);

                    for (int ti = startIdx; ti <= endIdx && ti < sqlFragment.ScriptTokenStream.Count; ti++)
                    {
                        var tok = sqlFragment.ScriptTokenStream[ti];
                        if (tok.TokenType == TSqlTokenType.WhiteSpace && tok.Column == 1)
                            tok.Text = RemoveOneIndent(tok.Text, indentString);
                    }
                }

            }

            // special case #9 - break DECLARE variables per line
            if (formatSettings.breakVariableDefinitionsPerLine)
            {
                var tokens = sqlFragment.ScriptTokenStream;
                foreach (var decl in visitor.DeclareStatements)
                {
                    // only split if more than one variable
                    if (decl.Declarations.Count > 1)
                    {
                        // figure out how much indent DECLARE already has
                        string baseIndent = "";
                        int wsBefore = decl.FirstTokenIndex - 1;
                        if (wsBefore >= 0 && tokens[wsBefore].TokenType == TSqlTokenType.WhiteSpace)
                        {
                            var txt = tokens[wsBefore].Text;
                            int lastNl = txt.LastIndexOf("\r\n");
                            baseIndent = lastNl >= 0
                                ? txt.Substring(lastNl + 2)
                                : txt;
                        }

                        // 1) break after the first declaration
                        int endOfFirst = decl.Declarations[0].LastTokenIndex + 1;
                        if (endOfFirst < tokens.Count && tokens[endOfFirst].TokenType == TSqlTokenType.WhiteSpace)
                        {
                            tokens[endOfFirst].Text = "\r\n"
                                                     + baseIndent
                                                     + "\t";
                        }

                        // 2) for each comma in the DECLARE, break+indent
                        for (int i = decl.FirstTokenIndex; i <= decl.LastTokenIndex; i++)
                        {
                            if (tokens[i].TokenType == TSqlTokenType.Comma)
                            {
                                int afterComma = i + 1;
                                if (afterComma < tokens.Count
                                    && tokens[afterComma].TokenType == TSqlTokenType.WhiteSpace)
                                {
                                    tokens[afterComma].Text = " "       // keep a space
                                                               + "\r\n"
                                                               + baseIndent
                                                               + "\t";
                                }
                            }
                        }
                    }
                }
            }

            // special case #10 - split sproc definition parameters per line + add one tab
            if (formatSettings.breakSprocDefinitionParametersPerLine)
            {
                var tokens = sqlFragment.ScriptTokenStream;
                const string indent = "\t";

                // helper to process any one proc‐definition
                void SplitParams(IList<ProcedureParameter> parameters)
                {
                    if (parameters.Count <= 1)
                        return;

                    // break before the very first param
                    int wsIdx = parameters[0].FirstTokenIndex - 1;
                    if (wsIdx >= 0 && tokens[wsIdx].TokenType == TSqlTokenType.WhiteSpace)
                        tokens[wsIdx].Text = "\r\n" + indent;

                    // now break at every comma
                    int start = parameters.First().FirstTokenIndex;
                    int end = parameters.Last().LastTokenIndex;
                    for (int i = start; i <= end; i++)
                    {
                        if (tokens[i].TokenType == TSqlTokenType.Comma)
                        {
                            int after = i + 1;
                            if (after < tokens.Count
                             && tokens[after].TokenType == TSqlTokenType.WhiteSpace)
                                tokens[after].Text = " "      // keep one space
                                                   + "\r\n"
                                                   + indent;
                        }
                    }
                }

                // apply to all three collections
                foreach (var proc in visitor.SprocDefinitionsCreate)
                    SplitParams(proc.Parameters);
                foreach (var proc in visitor.SprocDefinitionsAlter)
                    SplitParams(proc.Parameters);
                foreach (var proc in visitor.SprocDefinitionsCreateAlter)
                    SplitParams(proc.Parameters);

            }

            // special case #11 - split SELECT fields after TOP, fixed indent up to FROM
            // TODO - can't get this right, so disabled for now
            if (formatSettings.breakSelectFieldsAfterTopAndUnindent)
            {
                //var tokens = sqlFragment.ScriptTokenStream;
                //const string indent = "\t";

                //foreach (var qs in visitor.SelectsWithTop)
                //{
                //    // 1) break right after TOP(...) into "\r\n\t"
                //    int splitIdx = qs.TopRowFilter.LastTokenIndex + 1;
                //    if (splitIdx < tokens.Count
                //     && tokens[splitIdx].TokenType == TSqlTokenType.WhiteSpace)
                //    {
                //        tokens[splitIdx].Text = "\r\n" + indent;
                //    }

                //    //// 2) for every select‐element after the first, break before it into "\r\n\t"
                //    //for (int i = 1; i < qs.SelectElements.Count; i++)
                //    //{
                //    //    var elem = qs.SelectElements[i];
                //    //    int wsIdx = elem.FirstTokenIndex - 1;
                //    //    if (wsIdx >= 0
                //    //     && tokens[wsIdx].TokenType == TSqlTokenType.WhiteSpace)
                //    //    {
                //    //        tokens[wsIdx].Text = "\r\n" + indent;
                //    //    }
                //    //}

                //    //// 3) unindent FROM back to column 1
                //    //int wsBeforeFrom = qs.FromClause.FirstTokenIndex - 1;
                //    //if (wsBeforeFrom >= 0
                //    // && tokens[wsBeforeFrom].TokenType == TSqlTokenType.WhiteSpace)
                //    //{
                //    //    tokens[wsBeforeFrom].Text = "\r\n";
                //    //}
                //}
            }

            // return full recompiled result
            StringBuilder sqlText = new StringBuilder();
            foreach (var Token in sqlFragment.ScriptTokenStream)
            {
                sqlText.Append(Token.Text);
            }

            return sqlText.ToString();
        }

        // Helper: for a whitespace token that looks like "\r\n    …",
        // remove exactly one instance of indentString after each newline.
        private static string RemoveOneIndent(string wsText, string indentString)
        {
            var lines = wsText.Split(new[] { "\r\n" }, StringSplitOptions.None);
            // if it’s just spaces/tabs, drop one indent if present
            if (lines.Length == 1)
            {
                return lines[0].StartsWith(indentString)
                    ? lines[0].Substring(indentString.Length)
                    : lines[0];
            }
            // otherwise for each line after the first, drop one indent if present
            for (int i = 1; i < lines.Length; i++)
            {
                if (lines[i].StartsWith(indentString))
                    lines[i] = lines[i].Substring(indentString.Length);
            }
            return string.Join("\r\n", lines);
        }


    }
}
