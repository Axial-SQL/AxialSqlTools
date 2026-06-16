using Microsoft.SqlServer.TransactSql.ScriptDom;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
            public List<StatementWithCtesAndXmlNamespaces> StatementsWithCtes = new List<StatementWithCtesAndXmlNamespaces>();

            private void TrackStatementWithCtes(StatementWithCtesAndXmlNamespaces node)
            {
                if (node.WithCtesAndXmlNamespaces?.CommonTableExpressions?.Count > 0)
                {
                    StatementsWithCtes.Add(node);
                }
            }

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

            public override void ExplicitVisit(SelectStatement node)
            {
                base.ExplicitVisit(node);
                TrackStatementWithCtes(node);
            }

            public override void ExplicitVisit(InsertStatement node)
            {
                base.ExplicitVisit(node);
                TrackStatementWithCtes(node);
            }

            public override void ExplicitVisit(UpdateStatement node)
            {
                base.ExplicitVisit(node);
                TrackStatementWithCtes(node);
            }

            public override void ExplicitVisit(DeleteStatement node)
            {
                base.ExplicitVisit(node);
                TrackStatementWithCtes(node);
            }

            public override void ExplicitVisit(MergeStatement node)
            {
                base.ExplicitVisit(node);
                TrackStatementWithCtes(node);
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
            var formatSettings = SettingsManager.GetTSqlCodeFormatSettings();
            if (settingsOverride != null)
            {
                formatSettings = settingsOverride;
            }

            if (!HasTransformFormattingEnabled(formatSettings))
            {
                return oldCode;
            }

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

            Sql170ScriptGenerator gen = new Sql170ScriptGenerator();
            gen.Options.AlignClauseBodies = false;
            gen.Options.AlignColumnDefinitionFields = false;
            gen.Options.MultilineSelectElementsList = formatSettings.breakSelectElementsPerLine;
            gen.Options.NewLineBeforeFromClause = formatSettings.breakSelectElementsPerLine;
            gen.Options.NewLineBeforeWhereClause = formatSettings.breakSelectElementsPerLine;
            gen.Options.NewLineBeforeJoinClause = formatSettings.breakSelectElementsPerLine;
            gen.Options.NewLineBeforeCloseParenthesisInMultilineList = formatSettings.formatTableDefinitionsMultiline || formatSettings.breakSelectElementsPerLine;
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

            resultCode = PreserveLeadingCommaIndentAfterInlineComments(oldCode, resultCode);
            resultCode = PreserveSourceTerminalShape(oldCode, resultCode);
            return resultCode;

        }

        private static string PreserveLeadingCommaIndentAfterInlineComments(string oldCode, string resultCode)
        {
            if (string.IsNullOrEmpty(oldCode) || string.IsNullOrEmpty(resultCode))
            {
                return resultCode;
            }

            var originalLines = NormalizeLineEndings(oldCode).Split(new[] { "\r\n" }, StringSplitOptions.None);
            var resultLines = NormalizeLineEndings(resultCode).Split(new[] { "\r\n" }, StringSplitOptions.None);
            int searchStart = 0;

            for (int i = 0; i < originalLines.Length - 1; i++)
            {
                string originalCommentLine = originalLines[i];
                string originalNextLine = originalLines[i + 1];
                if (!HasInlineComment(originalCommentLine) || !IsLeadingCommaLine(originalNextLine))
                {
                    continue;
                }

                string originalIndent = LeadingWhitespace(originalNextLine);
                string commentLineKey = originalCommentLine.TrimEnd();
                for (int j = searchStart; j < resultLines.Length - 1; j++)
                {
                    if (resultLines[j].TrimEnd() != commentLineKey || !IsLeadingCommaLine(resultLines[j + 1]))
                    {
                        continue;
                    }

                    resultLines[j + 1] = originalIndent + resultLines[j + 1].TrimStart();
                    searchStart = j + 1;
                    break;
                }
            }

            string result = string.Join("\r\n", resultLines);
            if (!EndsWithLineBreak(resultCode))
            {
                result = result.TrimEnd('\r', '\n');
            }

            return result;
        }

        private static bool HasInlineComment(string line)
        {
            int commentIndex = (line ?? string.Empty).IndexOf("--", StringComparison.Ordinal);
            return commentIndex > 0 && !string.IsNullOrWhiteSpace(line.Substring(0, commentIndex));
        }

        private static bool IsLeadingCommaLine(string line)
        {
            return Regex.IsMatch(line ?? string.Empty, @"^\s*,");
        }

        private static string LeadingWhitespace(string line)
        {
            return Regex.Match(line ?? string.Empty, @"^\s*").Value;
        }

        private static string PreserveSourceTerminalShape(string oldCode, string resultCode)
        {
            if (string.IsNullOrEmpty(resultCode))
            {
                return resultCode;
            }

            string result = resultCode;
            if (!LastSignificantCharacterIs(oldCode, ';') && LastSignificantCharacterIs(result, ';'))
            {
                int semicolonIndex = LastSignificantCharacterIndex(result);
                result = result.Remove(semicolonIndex, 1);
            }

            result = MatchTrailingLineBreaks(oldCode, result);

            return result;
        }

        private static string MatchTrailingLineBreaks(string oldCode, string resultCode)
        {
            int oldTrailingLineBreaks = CountTrailingLineBreaks(oldCode);
            string result = (resultCode ?? string.Empty).TrimEnd('\r', '\n');
            string lineEnding = DetectLineEnding(oldCode);
            for (int i = 0; i < oldTrailingLineBreaks; i++)
            {
                result += lineEnding;
            }

            return result;
        }

        private static int CountTrailingLineBreaks(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            int count = 0;
            int i = text.Length;
            while (i > 0)
            {
                if (text[i - 1] == '\n')
                {
                    count++;
                    i--;
                    if (i > 0 && text[i - 1] == '\r')
                    {
                        i--;
                    }
                }
                else if (text[i - 1] == '\r')
                {
                    count++;
                    i--;
                }
                else
                {
                    break;
                }
            }

            return count;
        }

        private static string DetectLineEnding(string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                int crlf = text.IndexOf("\r\n", StringComparison.Ordinal);
                if (crlf >= 0)
                {
                    return "\r\n";
                }

                if (text.IndexOf('\n') >= 0)
                {
                    return "\n";
                }

                if (text.IndexOf('\r') >= 0)
                {
                    return "\r";
                }
            }

            return Environment.NewLine;
        }

        private static bool LastSignificantCharacterIs(string text, char expected)
        {
            int index = LastSignificantCharacterIndex(text);
            return index >= 0 && text[index] == expected;
        }

        private static int LastSignificantCharacterIndex(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return -1;
            }

            for (int i = text.Length - 1; i >= 0; i--)
            {
                if (!char.IsWhiteSpace(text[i]))
                {
                    return i;
                }
            }

            return -1;
        }

        private static bool EndsWithLineBreak(string text)
        {
            return !string.IsNullOrEmpty(text)
                && (text.EndsWith("\r\n", StringComparison.Ordinal)
                    || text.EndsWith("\n", StringComparison.Ordinal)
                    || text.EndsWith("\r", StringComparison.Ordinal));
        }

        private static bool HasTransformFormattingEnabled(SettingsManager.TSqlCodeFormatSettings formatSettings)
        {
            if (formatSettings == null)
            {
                return false;
            }

            return formatSettings.removeNewLineAfterJoin
                || formatSettings.addTabAfterJoinOn
                || formatSettings.moveCrossJoinToNewLine
                || formatSettings.formatCaseAsMultiline
                || formatSettings.addNewLineBetweenStatementsInBlocks
                || formatSettings.breakSprocParametersPerLine
                || formatSettings.uppercaseBuiltInFunctions
                || formatSettings.unindentBeginEndBlocks
                || formatSettings.breakVariableDefinitionsPerLine
                || formatSettings.breakSprocDefinitionParametersPerLine
                || formatSettings.breakSelectFieldsAfterTopAndUnindent
                || formatSettings.leadingCommas
                || formatSettings.semicolonBeforeCte
                || formatSettings.breakSelectElementsPerLine
                || formatSettings.useAssignmentAliases
                || formatSettings.omitAsForTableAliases
                || formatSettings.omitAsInDeclare
                || formatSettings.formatTableDefinitionsMultiline
                || formatSettings.prefixUnicodeStrings
                || formatSettings.removeSemicolonsFromDeclare;
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

            //special case #3 - CROSS/OUTER JOIN/APPLY should be on the new line
            if (formatSettings.moveCrossJoinToNewLine)
            {
                foreach (UnqualifiedJoin CrossJoin in visitor.UnqualifiedJoins)
                {
                    MoveJoinKeywordToNewLine(
                        sqlFragment,
                        CrossJoin.SecondTableReference.FirstTokenIndex,
                        CrossJoin.StartColumn,
                        TSqlTokenType.Cross,
                        TSqlTokenType.Outer);
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

            // special case #12 - place commas at the start of continuation lines
            if (formatSettings.leadingCommas)
            {
                ApplyLeadingCommas(sqlFragment.ScriptTokenStream);
            }

            // special case #13 - add a defensive semicolon immediately before CTE WITH
            if (formatSettings.semicolonBeforeCte)
            {
                ApplySemicolonBeforeCtes(sqlFragment.ScriptTokenStream, visitor.StatementsWithCtes);
            }

            // return full recompiled result
            StringBuilder sqlText = new StringBuilder();
            foreach (var Token in sqlFragment.ScriptTokenStream)
            {
                sqlText.Append(Token.Text);
            }

            return ApplyTextFormat(sqlText.ToString(), formatSettings);
        }

        private static string ApplyTextFormat(string sql, SettingsManager.TSqlCodeFormatSettings formatSettings)
        {
            string result = NormalizeLineEndings(sql).Replace("\t", "    ").TrimStart('\r', '\n');

            if (formatSettings.omitAsInDeclare)
            {
                result = RemoveAsInDeclareStatements(result);
            }

            if (formatSettings.omitAsForTableAliases)
            {
                result = RemoveAsForTableAliases(result);
            }

            if (formatSettings.formatTableDefinitionsMultiline)
            {
                result = FormatTableVariableDefinitions(result);
            }

            if (formatSettings.breakSelectElementsPerLine)
            {
                result = FormatSelectLists(result, formatSettings.leadingCommas);
                result = FormatCteSelectOpen(result);
            }

            if (formatSettings.leadingCommas)
            {
                result = NormalizeLeadingCommaLines(result);
            }

            if (formatSettings.useAssignmentAliases)
            {
                result = ApplyAssignmentAliases(result);
            }

            if (formatSettings.prefixUnicodeStrings)
            {
                result = PrefixUnicodeStringLiterals(result);
            }

            if (formatSettings.removeSemicolonsFromDeclare)
            {
                result = RemoveDeclareSemicolons(result);
            }

            if (formatSettings.leadingCommas)
            {
                result = NormalizeLeadingCommaLines(result);
            }

            if (formatSettings.breakSelectElementsPerLine)
            {
                result = FormatCteSelectOpen(result);
                result = NormalizeCteJoinIndent(result);
                result = FormatExistingMultilineSelects(result, formatSettings.leadingCommas);
            }

            result = NormalizeInlineCommentSpacing(result);
            result = TrimTrailingSpaces(result);

            return NormalizeLineEndings(result);
        }

        private static string RemoveAsInDeclareStatements(string sql)
        {
            return Regex.Replace(
                sql,
                @"(?im)^(\s*(?:DECLARE\s+)?[,]?@\w+)\s+AS\s+",
                "$1 ");
        }

        private static string RemoveAsForTableAliases(string sql)
        {
            return Regex.Replace(
                sql,
                @"(?i)\b(FROM|JOIN|APPLY)\s+(@?\[?[\w.]+\]?)\s+AS\s+(\[?\w+\]?)",
                "$1 $2 $3");
        }

        private static string FormatTableVariableDefinitions(string sql)
        {
            var result = new StringBuilder();
            int position = 0;
            var regex = new Regex(@"(?im)^(?<indent>[ \t]*)DECLARE\s+(?<name>@\w+)\s+TABLE\s*\(");

            foreach (Match match in regex.Matches(sql))
            {
                if (match.Index < position)
                {
                    continue;
                }

                int openParen = sql.IndexOf('(', match.Index + match.Length - 1);
                int closeParen = FindMatchingParen(sql, openParen);
                if (openParen < 0 || closeParen < 0)
                {
                    continue;
                }

                int end = closeParen + 1;
                if (end < sql.Length && sql[end] == ';')
                {
                    end++;
                }

                string indent = match.Groups["indent"].Value;
                string childIndent = indent + "    ";
                string body = Regex.Replace(sql.Substring(openParen + 1, closeParen - openParen - 1), @"\s+", " ").Trim();
                var columns = SplitTopLevelCommaList(body);
                var builder = new StringBuilder();
                builder.Append(indent).Append("DECLARE ").Append(match.Groups["name"].Value).Append(" TABLE (");

                for (int i = 0; i < columns.Count; i++)
                {
                    builder.Append("\r\n").Append(childIndent);
                    if (i > 0)
                    {
                        builder.Append(",");
                    }
                    builder.Append(NormalizeSpaces(columns[i]));
                }

                builder.Append("\r\n").Append(indent).Append(")");

                result.Append(sql.Substring(position, match.Index - position));
                result.Append(builder.ToString());
                position = end;
            }

            result.Append(sql.Substring(position));
            return result.ToString();
        }

        private static string ApplyAssignmentAliases(string sql)
        {
            string result = sql;

            result = Regex.Replace(
                result,
                @"(?ims)^(?<indent>\s*)(?<comma>,?)(?<post>\s*)CASE\r?\n(?<body>.*?)^(?<endindent>\s*)END\s+AS\s+(?<alias>\w+)\s*$",
                match =>
                {
                    string indent = EffectiveCommaIndent(match);
                    string comma = string.IsNullOrEmpty(match.Groups["comma"].Value) ? string.Empty : ",";
                    string body = ReindentCaseBody(match.Groups["body"].Value, indent + "        ");
                    return $"{indent}{comma}{match.Groups["alias"].Value} =\r\n{indent}    CASE\r\n{body}\r\n{indent}    END";
                });

            result = Regex.Replace(
                result,
                @"(?im)^(?<indent>\s*)(?<comma>,?)(?<post>\s*)(?<expr>(?:[A-Z_][\w.]*\s*)?\([^\r\n]*\)|[\w.]+)\s+AS\s+(?<alias>\w+)\s*$",
                match =>
                {
                    string indent = EffectiveCommaIndent(match);
                    string comma = string.IsNullOrEmpty(match.Groups["comma"].Value) ? string.Empty : ",";
                    return $"{indent}{comma}{match.Groups["alias"].Value} = {match.Groups["expr"].Value.Trim()}";
                });

            return result;
        }

        private static string FormatSelectLists(string sql, bool leadingCommas)
        {
            return Regex.Replace(
                sql,
                @"(?im)^(?<indent>[ \t]*)SELECT(?<top>\s+TOP\s+\(?\d+\)?)?\s+(?<items>.+?(?:--.*)?)$",
                match =>
                {
                    string itemsText = match.Groups["items"].Value;
                    if (!itemsText.Contains(",") || Regex.IsMatch(itemsText, @"(?i)\bFROM\b"))
                    {
                        return match.Value;
                    }

                    string indent = match.Groups["indent"].Value;
                    string childIndent = indent + "    ";
                    var items = SplitTopLevelCommaList(itemsText);
                    if (items.Count <= 1)
                    {
                        return match.Value;
                    }

                    var builder = new StringBuilder();
                    builder.Append(indent).Append("SELECT").Append(match.Groups["top"].Value);
                    for (int i = 0; i < items.Count; i++)
                    {
                        builder.Append("\r\n").Append(childIndent);
                        if (i > 0 && leadingCommas)
                        {
                            builder.Append(",");
                        }
                        builder.Append(items[i].Trim());
                        if (i < items.Count - 1 && !leadingCommas)
                        {
                            builder.Append(",");
                        }
                    }

                    return builder.ToString();
                });
        }

        private static string FormatCteSelectOpen(string sql)
        {
            string result = Regex.Replace(
                sql,
                @"(?im)^(?<indent>[ \t]*);?WITH\s+(?<name>\w+)[ \t]*\r?\n[ \t]*AS\s*\(SELECT(?<top>\s+TOP\s+\(?\d+\)?)?\s+(?<first>[^\r\n,]+)",
                match =>
                {
                    string indent = match.Groups["indent"].Value;
                    return $"{indent};WITH {match.Groups["name"].Value} AS (\r\n{indent}    SELECT{match.Groups["top"].Value}\r\n{indent}        {match.Groups["first"].Value.Trim()}";
                });

            result = Regex.Replace(
                result,
                @"(?im)^(?<indent>[ \t]*)(?<join>INNER|LEFT|RIGHT|FULL)\s+JOIN\s*\r?\n[ \t]*(?<table>@?\w+\s+\w+)\s*$",
                "${indent}${join} JOIN ${table}");

            result = Regex.Replace(
                result,
                @"(?im)^(?<indent>[ \t]*)(?<line>CROSS\s+JOIN\s+@?\w+\s+\w+)\)$",
                "${indent}${line}\r\n${indent})");

            return result;
        }

        private static string NormalizeCteJoinIndent(string sql)
        {
            var lines = NormalizeLineEndings(sql).Split(new[] { "\r\n" }, StringSplitOptions.None);
            bool insideCte = false;
            string cteIndent = string.Empty;
            string fromIndent = string.Empty;
            var output = new List<string>();

            foreach (string originalLine in lines)
            {
                string line = originalLine;
                string trimmed = line.TrimStart();
                string currentIndent = Regex.Match(line, @"^[ \t]*").Value;

                if (Regex.IsMatch(trimmed, @"^;WITH\b", RegexOptions.IgnoreCase))
                {
                    insideCte = true;
                    cteIndent = currentIndent;
                    fromIndent = string.Empty;
                    output.Add(line);
                    continue;
                }

                if (insideCte && Regex.IsMatch(trimmed, @"^FROM\b", RegexOptions.IgnoreCase))
                {
                    fromIndent = currentIndent;
                    output.Add(line);
                    continue;
                }

                if (insideCte && !string.IsNullOrEmpty(fromIndent))
                {
                    if (Regex.IsMatch(trimmed, @"^(INNER|LEFT|RIGHT|FULL|CROSS)\s+JOIN\b", RegexOptions.IgnoreCase))
                    {
                        bool closesCte = trimmed.EndsWith(")", StringComparison.Ordinal);
                        string joinText = closesCte ? trimmed.Substring(0, trimmed.Length - 1).TrimEnd() : trimmed;
                        output.Add(fromIndent + joinText);
                        if (closesCte)
                        {
                            output.Add(cteIndent + ")");
                            insideCte = false;
                            fromIndent = string.Empty;
                        }

                        continue;
                    }

                    if (Regex.IsMatch(trimmed, @"^ON\b", RegexOptions.IgnoreCase))
                    {
                        output.Add(fromIndent + "    " + trimmed);
                        continue;
                    }
                }

                if (insideCte && trimmed == ")")
                {
                    output.Add(cteIndent + ")");
                    insideCte = false;
                    fromIndent = string.Empty;
                    continue;
                }

                output.Add(line);
            }

            return string.Join("\r\n", output);
        }

        private static string FormatExistingMultilineSelects(string sql, bool leadingCommas)
        {
            return Regex.Replace(
                sql,
                @"(?im)^(?<indent>[ \t]*)SELECT\s+(?<first>[^\r\n,]+)\r?\n(?<items>(?:[ \t]*,[^\r\n]*(?:\r?\n|$)){1,})",
                match =>
                {
                    string indent = match.Groups["indent"].Value;
                    string childIndent = indent + "    ";
                    var items = new List<string> { match.Groups["first"].Value.Trim() };
                    foreach (Match itemMatch in Regex.Matches(match.Groups["items"].Value, @"(?m)^[ \t]*,(?<item>[^\r\n]*)"))
                    {
                        items.Add(itemMatch.Groups["item"].Value.Trim());
                    }

                    if (items.Count <= 1)
                    {
                        return match.Value;
                    }

                    var builder = new StringBuilder();
                    builder.Append(indent).Append("SELECT");
                    for (int i = 0; i < items.Count; i++)
                    {
                        builder.Append("\r\n").Append(childIndent);
                        if (i > 0 && leadingCommas)
                        {
                            builder.Append(",");
                        }

                        builder.Append(items[i]);
                        if (i < items.Count - 1 && !leadingCommas)
                        {
                            builder.Append(",");
                        }
                    }

                    return builder.ToString();
                });
        }

        private static string NormalizeLeadingCommaLines(string sql)
        {
            var lines = NormalizeLineEndings(sql).Split(new[] { "\r\n" }, StringSplitOptions.None);
            string previousItemIndent = string.Empty;

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                Match commaMatch = Regex.Match(line, @"^(?<indent>\s*),(?<space>\s*)(?<text>.*)$");
                if (commaMatch.Success)
                {
                    string candidateIndent = commaMatch.Groups["indent"].Value;
                    string indent = string.IsNullOrEmpty(candidateIndent)
                        ? previousItemIndent
                        : candidateIndent;
                    if (!string.IsNullOrEmpty(previousItemIndent)
                        && candidateIndent.Length > previousItemIndent.Length + 4)
                    {
                        indent = previousItemIndent;
                    }

                    lines[i] = indent + "," + commaMatch.Groups["text"].Value.TrimStart();
                    previousItemIndent = indent;
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(line)
                    && !Regex.IsMatch(line.TrimStart(), @"^(SELECT|FROM|WHERE|INNER|LEFT|RIGHT|FULL|CROSS|ON|ELSE|END|\)|;WITH)\b", RegexOptions.IgnoreCase))
                {
                    previousItemIndent = Regex.Match(line, @"^\s*").Value;
                }
            }

            return string.Join("\r\n", lines);
        }

        private static string EffectiveCommaIndent(Match match)
        {
            string indent = match.Groups["indent"].Value;
            if (string.IsNullOrEmpty(indent) && !string.IsNullOrEmpty(match.Groups["comma"].Value))
            {
                indent = match.Groups["post"].Value;
            }

            return indent;
        }

        private static int FindMatchingParen(string text, int openParen)
        {
            if (openParen < 0)
            {
                return -1;
            }

            int depth = 0;
            for (int i = openParen; i < text.Length; i++)
            {
                if (text[i] == '(')
                {
                    depth++;
                }
                else if (text[i] == ')')
                {
                    depth--;
                    if (depth == 0)
                    {
                        return i;
                    }
                }
            }

            return -1;
        }

        private static string ReindentCaseBody(string body, string indent)
        {
            var sourceLines = NormalizeLineEndings(body).Split(new[] { "\r\n" }, StringSplitOptions.None)
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .Select(line => line.Trim())
                .ToList();

            var lines = new List<string>();
            for (int i = 0; i < sourceLines.Count; i++)
            {
                string line = sourceLines[i];
                if (line.StartsWith("WHEN ", StringComparison.OrdinalIgnoreCase)
                    && i + 1 < sourceLines.Count
                    && sourceLines[i + 1].StartsWith("THEN ", StringComparison.OrdinalIgnoreCase))
                {
                    lines.Add(indent + line + " " + sourceLines[i + 1]);
                    i++;
                    continue;
                }

                lines.Add(indent + line);
            }

            return string.Join("\r\n", lines);
        }

        private static string NormalizeInlineCommentSpacing(string sql)
        {
            return Regex.Replace(
                sql,
                @"(?m)(--[^\r\n]*)\r\n[ \t]*\r\n([ \t]*(?:FROM|WHERE|GROUP|ORDER|HAVING)\b)",
                "$1\r\n$2");
        }

        private static string TrimTrailingSpaces(string sql)
        {
            return Regex.Replace(sql, @"(?m)[ \t]+(?=\r?$)", string.Empty);
        }

        private static string PrefixUnicodeStringLiterals(string sql)
        {
            var builder = new StringBuilder();
            for (int i = 0; i < sql.Length; i++)
            {
                if (i + 1 < sql.Length && sql[i] == '-' && sql[i + 1] == '-')
                {
                    int end = sql.IndexOf('\n', i + 2);
                    if (end < 0)
                    {
                        builder.Append(sql.Substring(i));
                        break;
                    }

                    builder.Append(sql.Substring(i, end - i + 1));
                    i = end;
                    continue;
                }

                if (i + 1 < sql.Length && sql[i] == '/' && sql[i + 1] == '*')
                {
                    int end = sql.IndexOf("*/", i + 2, StringComparison.Ordinal);
                    if (end < 0)
                    {
                        builder.Append(sql.Substring(i));
                        break;
                    }

                    builder.Append(sql.Substring(i, end - i + 2));
                    i = end + 1;
                    continue;
                }

                if (sql[i] == '\'')
                {
                    if (i == 0 || (sql[i - 1] != 'N' && sql[i - 1] != 'n'))
                    {
                        builder.Append('N');
                    }

                    builder.Append(sql[i]);
                    while (++i < sql.Length)
                    {
                        builder.Append(sql[i]);
                        if (sql[i] == '\'')
                        {
                            if (i + 1 < sql.Length && sql[i + 1] == '\'')
                            {
                                builder.Append(sql[++i]);
                                continue;
                            }

                            break;
                        }
                    }

                    continue;
                }

                builder.Append(sql[i]);
            }

            return builder.ToString();
        }

        private static string RemoveDeclareSemicolons(string sql)
        {
            return Regex.Replace(sql, @"(?im)^(\s*(?:DECLARE\b|,@).*?);(\s*)$", "$1$2");
        }

        private static List<string> SplitTopLevelCommaList(string text)
        {
            var items = new List<string>();
            int depth = 0;
            int start = 0;

            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (ch == '(')
                {
                    depth++;
                }
                else if (ch == ')' && depth > 0)
                {
                    depth--;
                }
                else if (ch == ',' && depth == 0)
                {
                    items.Add(text.Substring(start, i - start).Trim());
                    start = i + 1;
                }
            }

            if (start < text.Length)
            {
                items.Add(text.Substring(start).Trim());
            }

            return items;
        }

        private static string NormalizeSpaces(string text)
        {
            return Regex.Replace(text ?? string.Empty, @"\s+", " ").Trim();
        }

        private static string NormalizeLineEndings(string text)
        {
            return (text ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "\r\n");
        }

        private static void ApplyLeadingCommas(IList<TSqlParserToken> tokens)
        {
            for (int i = 0; i < tokens.Count - 1; i++)
            {
                if (tokens[i].TokenType != TSqlTokenType.Comma)
                {
                    continue;
                }

                int wsAfterComma = i + 1;
                if (tokens[wsAfterComma].TokenType != TSqlTokenType.WhiteSpace
                    || !ContainsNewLine(tokens[wsAfterComma].Text))
                {
                    continue;
                }

                tokens[i].Text = string.Empty;
                tokens[wsAfterComma].Text = MoveCommaToContinuationLine(tokens[wsAfterComma].Text);
            }
        }

        private static void ApplySemicolonBeforeCtes(
            IList<TSqlParserToken> tokens,
            IList<StatementWithCtesAndXmlNamespaces> statementsWithCtes)
        {
            foreach (var statement in statementsWithCtes)
            {
                if (statement.FirstTokenIndex < 0 || statement.FirstTokenIndex >= tokens.Count)
                {
                    continue;
                }

                int cteTokenIndex = statement.FirstTokenIndex;
                int previousSignificantTokenIndex = cteTokenIndex - 1;
                while (previousSignificantTokenIndex >= 0
                    && tokens[previousSignificantTokenIndex].TokenType == TSqlTokenType.WhiteSpace)
                {
                    previousSignificantTokenIndex--;
                }

                int whitespaceBeforeCte = cteTokenIndex - 1;
                if (previousSignificantTokenIndex >= 0
                    && tokens[previousSignificantTokenIndex].TokenType == TSqlTokenType.Semicolon)
                {
                    if (previousSignificantTokenIndex == cteTokenIndex - 1)
                    {
                        continue;
                    }

                    tokens[previousSignificantTokenIndex].Text = string.Empty;
                }

                if (whitespaceBeforeCte >= 0
                    && tokens[whitespaceBeforeCte].TokenType == TSqlTokenType.WhiteSpace)
                {
                    tokens[whitespaceBeforeCte].Text = tokens[whitespaceBeforeCte].Text + ";";
                }
                else
                {
                    tokens[cteTokenIndex].Text = ";" + tokens[cteTokenIndex].Text;
                }
            }
        }

        private static bool ContainsNewLine(string text)
        {
            return text != null && (text.Contains("\r\n") || text.Contains("\n"));
        }

        private static string MoveCommaToContinuationLine(string whitespace)
        {
            if (string.IsNullOrEmpty(whitespace))
            {
                return ",";
            }

            int newlineIndex = whitespace.LastIndexOf("\r\n", StringComparison.Ordinal);
            int newlineLength = 2;

            if (newlineIndex < 0)
            {
                newlineIndex = whitespace.LastIndexOf("\n", StringComparison.Ordinal);
                newlineLength = 1;
            }

            if (newlineIndex < 0)
            {
                return "," + whitespace;
            }

            return whitespace.Substring(newlineIndex, newlineLength)
                + whitespace.Substring(newlineIndex + newlineLength)
                + ",";
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

        private static void MoveJoinKeywordToNewLine(
            TSqlFragment sqlFragment,
            int startTokenIndex,
            int startColumn,
            params TSqlTokenType[] tokenTypes)
        {
            var tokenSet = new HashSet<TSqlTokenType>(tokenTypes);
            int nextTokenNumber = startTokenIndex;

            while (true)
            {
                TSqlParserToken nextToken = sqlFragment.ScriptTokenStream[nextTokenNumber];

                if (tokenSet.Contains(nextToken.TokenType))
                {
                    TSqlParserToken previousToken = sqlFragment.ScriptTokenStream[nextTokenNumber - 1];
                    if (previousToken.TokenType == TSqlTokenType.WhiteSpace)
                        previousToken.Text = "\r\n" + new string(' ', startColumn - 1);
                    break;
                }

                nextTokenNumber -= 1;

                //just in case
                if (nextTokenNumber < 0)
                    break;
            }
        }


    }
}
