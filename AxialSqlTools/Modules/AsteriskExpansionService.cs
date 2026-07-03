using Microsoft.Data.SqlClient;
using Microsoft.SqlServer.TransactSql.ScriptDom;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;

namespace AxialSqlTools
{
    public static class AsteriskExpansionService
    {
        private sealed class SelectInfo
        {
            public string MetadataStatementText { get; set; }
            public int StatementStartOffset { get; set; }
            public Dictionary<string, string> TableQualifiersByName { get; set; }
            public List<TableInfo> Tables { get; set; }
        }

        private sealed class TableInfo
        {
            public string SchemaName { get; set; }
            public string TableName { get; set; }
            public string Qualifier { get; set; }
        }

        private sealed class ColumnInfo
        {
            public string Name { get; set; }
            public string SourceTableName { get; set; }
            public string Qualifier { get; set; } = string.Empty;
        }

        public static bool TryExpand(IVsTextView textView)
        {
            try
            {
                if (textView == null || textView.GetBuffer(out IVsTextLines textLines) != VSConstants.S_OK)
                    return false;

                textView.GetCaretPos(out int line, out int column);
                if (!TryGetAsteriskAtCaret(textLines, line, column, out int starColumn, out string qualifier))
                    return false;

                if (!TryGetFullText(textLines, out string fullText))
                    return false;

                int starOffset = GetAbsoluteOffset(fullText, line, starColumn);
                if (!TryGetCurrentSelectStatement(fullText, line + 1, column + 1, starOffset, out SelectInfo selectInfo))
                    return false;

                string tempTableSetup = GetPriorTempTableCreateStatements(fullText, selectInfo.StatementStartOffset);
                List<ColumnInfo> columns = GetResultColumns(selectInfo, tempTableSetup, qualifier);
                if (columns.Count == 0)
                    return false;

                string replacement = BuildColumnList(columns, starColumn, qualifier);
                ReplaceText(textLines, line, starColumn, column, replacement);
                SetCaretPosition(textView, line, starColumn, replacement, replacement.Length);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string BuildColumnList(IList<ColumnInfo> columns, int starColumn, string qualifier)
        {
            var escaped = new List<string>();
            foreach (ColumnInfo column in columns)
            {
                if (!string.IsNullOrWhiteSpace(column.Name))
                    escaped.Add(column.Qualifier + "[" + column.Name.Replace("]", "]]") + "]");
            }

            if (escaped.Count == 0)
                return string.Empty;

            string indent = new string(' ', Math.Max(0, starColumn - qualifier.Length));
            string separator = "," + Environment.NewLine + indent + qualifier;
            return string.Join(separator, escaped);
        }

        private static bool TryGetAsteriskAtCaret(IVsTextLines textLines, int line, int column, out int starColumn, out string qualifier)
        {
            starColumn = -1;
            qualifier = string.Empty;

            if (column <= 0)
                return false;

            textLines.GetLengthOfLine(line, out int lineLength);
            if (column > lineLength)
                return false;

            textLines.GetLineText(line, 0, line, lineLength, out string lineText);
            starColumn = column - 1;
            if (string.IsNullOrEmpty(lineText) || starColumn >= lineText.Length || lineText[starColumn] != '*')
                return false;

            qualifier = GetQualifierBeforeAsterisk(lineText, starColumn);
            return true;
        }

        private static string GetQualifierBeforeAsterisk(string lineText, int starColumn)
        {
            if (starColumn < 2 || lineText[starColumn - 1] != '.')
                return string.Empty;

            int start = starColumn - 2;
            if (lineText[start] == ']')
            {
                while (start >= 0 && lineText[start] != '[')
                    start--;

                if (start < 0)
                    return string.Empty;
            }
            else
            {
                while (start >= 0 && (char.IsLetterOrDigit(lineText[start]) || lineText[start] == '_' || lineText[start] == '#'))
                    start--;

                start++;
            }

            return lineText.Substring(start, starColumn - start);
        }

        private static bool TryGetFullText(IVsTextLines textLines, out string fullText)
        {
            fullText = null;

            if (textLines.GetLastLineIndex(out int lastLine, out int lastIndex) != VSConstants.S_OK)
                return false;

            return textLines.GetLineText(0, 0, lastLine, lastIndex, out fullText) == VSConstants.S_OK
                && !string.IsNullOrWhiteSpace(fullText);
        }

        private static bool TryGetCurrentSelectStatement(string fullText, int cursorLine, int cursorColumn, int starOffset, out SelectInfo selectInfo)
        {
            selectInfo = null;

            var parser = new TSql170Parser(false);
            TSqlFragment fragment = parser.Parse(new StringReader(fullText), out IList<ParseError> _);
            if (!(fragment is TSqlScript script) || fragment.ScriptTokenStream == null)
                return false;

            foreach (TSqlBatch batch in script.Batches)
            {
                foreach (TSqlStatement statement in batch.Statements)
                {
                    if (!(statement is SelectStatement) || !ContainsCursor(fragment, statement, cursorLine, cursorColumn))
                        continue;

                    if (statement.StartOffset < 0 || statement.FragmentLength <= 0 || statement.StartOffset + statement.FragmentLength > fullText.Length)
                        return false;

                    return TryBuildSelectInfo(fullText, (SelectStatement)statement, starOffset, out selectInfo);
                }
            }

            return false;
        }

        private static bool TryBuildSelectInfo(string fullText, SelectStatement statement, int starOffset, out SelectInfo selectInfo)
        {
            selectInfo = null;

            var query = statement.QueryExpression as QuerySpecification;
            if (query == null || query.SelectElements == null || query.SelectElements.Count == 0)
                return false;

            SelectElement selectedStar = null;
            foreach (SelectElement element in query.SelectElements)
            {
                if (element is SelectStarExpression && ContainsOffset(element, starOffset))
                {
                    selectedStar = element;
                    break;
                }
            }

            if (selectedStar == null)
                return false;

            SelectElement firstElement = query.SelectElements[0];
            SelectElement lastElement = query.SelectElements[query.SelectElements.Count - 1];
            int selectListStart = firstElement.StartOffset;
            int selectListEnd = lastElement.StartOffset + lastElement.FragmentLength;

            if (selectListStart < 0 || selectListEnd <= selectListStart || selectListEnd > fullText.Length)
                return false;

            string statementText = fullText.Substring(statement.StartOffset, statement.FragmentLength);
            string starText = fullText.Substring(selectedStar.StartOffset, selectedStar.FragmentLength);
            int relativeSelectListStart = selectListStart - statement.StartOffset;
            int relativeSelectListEnd = selectListEnd - statement.StartOffset;

            selectInfo = new SelectInfo
            {
                StatementStartOffset = statement.StartOffset,
                MetadataStatementText = statementText.Substring(0, relativeSelectListStart)
                    + starText
                    + statementText.Substring(relativeSelectListEnd),
                TableQualifiersByName = GetTableQualifiers(query.FromClause),
                Tables = GetTables(query.FromClause)
            };
            return true;
        }

        private static bool ContainsOffset(TSqlFragment fragment, int offset)
        {
            return fragment.StartOffset <= offset && offset < fragment.StartOffset + fragment.FragmentLength;
        }

        private static bool ContainsCursor(TSqlFragment fragment, TSqlStatement statement, int cursorLine, int cursorColumn)
        {
            TSqlParserToken firstToken = fragment.ScriptTokenStream[statement.FirstTokenIndex];
            TSqlParserToken lastToken = fragment.ScriptTokenStream[statement.LastTokenIndex];
            int endColumn = lastToken.Column + lastToken.Text.Length;

            if (cursorLine < firstToken.Line || (cursorLine == firstToken.Line && cursorColumn < firstToken.Column))
                return false;

            if (cursorLine > lastToken.Line || (cursorLine == lastToken.Line && cursorColumn > endColumn))
                return false;

            return true;
        }

        private static string GetPriorTempTableCreateStatements(string fullText, int beforeOffset)
        {
            if (string.IsNullOrWhiteSpace(fullText) || beforeOffset <= 0)
                return string.Empty;

            var parser = new TSql170Parser(false);
            TSqlFragment fragment = parser.Parse(new StringReader(fullText), out IList<ParseError> _);
            if (!(fragment is TSqlScript script))
                return string.Empty;

            var statements = new List<string>();
            foreach (TSqlBatch batch in script.Batches)
            {
                foreach (TSqlStatement statement in batch.Statements)
                {
                    if (statement.StartOffset >= beforeOffset)
                        continue;

                    if (statement is CreateTableStatement createTable
                        && IsTempTable(createTable.SchemaObjectName)
                        && statement.StartOffset >= 0
                        && statement.FragmentLength > 0
                        && statement.StartOffset + statement.FragmentLength <= fullText.Length)
                    {
                        statements.Add(fullText.Substring(statement.StartOffset, statement.FragmentLength));
                    }
                }
            }

            return string.Join(Environment.NewLine, statements);
        }

        private static bool IsTempTable(SchemaObjectName name)
        {
            if (name == null || name.BaseIdentifier == null)
                return false;

            string value = name.BaseIdentifier.Value;
            return !string.IsNullOrWhiteSpace(value) && value.StartsWith("#", StringComparison.Ordinal);
        }

        private static List<ColumnInfo> GetResultColumns(SelectInfo selectInfo, string tempTableSetup, string typedQualifier)
        {
            var columns = new List<ColumnInfo>();
            var connectionInfo = ScriptFactoryAccess.GetCurrentConnectionInfo();
            if (connectionInfo == null || string.IsNullOrWhiteSpace(connectionInfo.FullConnectionString))
                return columns;

            using (var connection = new SqlConnection(connectionInfo.FullConnectionString))
            {
                connection.Open();

                if (!string.IsNullOrWhiteSpace(tempTableSetup))
                {
                    using (var setupCommand = connection.CreateCommand())
                    {
                        setupCommand.CommandText = tempTableSetup;
                        setupCommand.CommandTimeout = 5;
                        setupCommand.ExecuteNonQuery();
                    }
                }

                columns = GetColumnsFromTableReferences(connection, selectInfo.Tables, typedQualifier);
                if (columns.Count > 0)
                    return columns;

                return GetResultColumnsFromFmtOnly(connection, selectInfo.MetadataStatementText, selectInfo.TableQualifiersByName);
            }
        }

        private static List<ColumnInfo> GetColumnsFromTableReferences(SqlConnection connection, List<TableInfo> tables, string typedQualifier)
        {
            var result = new List<ColumnInfo>();
            if (tables == null || tables.Count == 0)
                return result;

            List<TableInfo> targetTables = tables;
            if (!string.IsNullOrWhiteSpace(typedQualifier))
            {
                string normalizedQualifier = typedQualifier.TrimEnd('.');
                targetTables = tables.FindAll(t => string.Equals(t.Qualifier.TrimEnd('.'), normalizedQualifier, StringComparison.OrdinalIgnoreCase));
                if (targetTables.Count == 0)
                    return result;
            }

            foreach (TableInfo table in targetTables)
            {
                List<string> tableColumns = GetColumnNamesForTable(connection, table);
                if (tableColumns.Count == 0)
                    return new List<ColumnInfo>();

                foreach (string columnName in tableColumns)
                {
                    result.Add(new ColumnInfo
                    {
                        Name = columnName,
                        SourceTableName = table.TableName,
                        Qualifier = string.Empty
                    });
                }
            }

            if (string.IsNullOrWhiteSpace(typedQualifier))
                ApplyDuplicateQualifiers(result, BuildQualifierMap(targetTables));

            return result;
        }

        private static List<string> GetColumnNamesForTable(SqlConnection connection, TableInfo table)
        {
            var result = new List<string>();
            if (table == null || string.IsNullOrWhiteSpace(table.TableName))
                return result;

            using (var command = connection.CreateCommand())
            {
                command.CommandTimeout = 5;
                command.Parameters.AddWithValue("@objectName", GetObjectNameForLookup(table));
                command.CommandText = table.TableName.StartsWith("#", StringComparison.Ordinal)
                    ? "SELECT c.[name] FROM tempdb.sys.columns c WHERE c.[object_id] = OBJECT_ID(@objectName) ORDER BY c.column_id;"
                    : "SELECT c.[name] FROM sys.columns c WHERE c.[object_id] = OBJECT_ID(@objectName) ORDER BY c.column_id;";

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        if (!reader.IsDBNull(0))
                            result.Add(reader.GetString(0));
                    }
                }
            }

            return result;
        }

        private static string GetObjectNameForLookup(TableInfo table)
        {
            if (table.TableName.StartsWith("#", StringComparison.Ordinal))
                return "tempdb.." + table.TableName;

            return string.IsNullOrWhiteSpace(table.SchemaName)
                ? table.TableName
                : table.SchemaName + "." + table.TableName;
        }

        private static Dictionary<string, string> BuildQualifierMap(List<TableInfo> tables)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (TableInfo table in tables)
            {
                if (!string.IsNullOrWhiteSpace(table.TableName))
                    result[table.TableName] = table.Qualifier;
            }

            return result;
        }

        private static List<ColumnInfo> GetResultColumnsFromFmtOnly(SqlConnection connection, string statementText, Dictionary<string, string> tableQualifiersByName)
        {
            var columns = new List<ColumnInfo>();

            using (var command = connection.CreateCommand())
            {
                // ponytail: FMTONLY is metadata-only fallback; ceiling is dynamic SQL/temp-table-heavy scripts. Upgrade path: richer ScriptDom + database metadata resolver.
                command.CommandText = "SET FMTONLY ON;" + Environment.NewLine
                    + statementText + Environment.NewLine
                    + "SET FMTONLY OFF;";
                command.CommandTimeout = 5;

                using (var reader = command.ExecuteReader())
                {
                    do
                    {
                        DataTable schema = reader.GetSchemaTable();
                        if (schema == null)
                            continue;

                        foreach (DataRow row in schema.Rows)
                        {
                            string name = row["ColumnName"] as string;
                            if (string.IsNullOrWhiteSpace(name))
                                continue;

                            columns.Add(new ColumnInfo
                            {
                                Name = name,
                                SourceTableName = GetSchemaString(row, "BaseTableName")
                            });
                        }

                        if (columns.Count > 0)
                        {
                            ApplyDuplicateQualifiers(columns, tableQualifiersByName);
                            return columns;
                        }
                    }
                    while (reader.NextResult());
                }
            }

            return columns;
        }

        private static string GetSchemaString(DataRow row, string columnName)
        {
            return row.Table.Columns.Contains(columnName) ? row[columnName] as string : null;
        }

        private static void ApplyDuplicateQualifiers(List<ColumnInfo> columns, Dictionary<string, string> tableQualifiersByName)
        {
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (ColumnInfo column in columns)
            {
                if (!counts.ContainsKey(column.Name))
                    counts[column.Name] = 0;

                counts[column.Name]++;
            }

            foreach (ColumnInfo column in columns)
            {
                if (counts[column.Name] <= 1)
                    continue;

                column.Qualifier = GetQualifier(column.SourceTableName, tableQualifiersByName);
            }
        }

        private static string GetQualifier(string sourceTableName, Dictionary<string, string> tableQualifiersByName)
        {
            if (string.IsNullOrWhiteSpace(sourceTableName))
                return string.Empty;

            if (tableQualifiersByName != null && tableQualifiersByName.TryGetValue(sourceTableName, out string qualifier))
                return qualifier;

            if (tableQualifiersByName != null)
            {
                foreach (var pair in tableQualifiersByName)
                {
                    if (pair.Key.StartsWith("#", StringComparison.Ordinal)
                        && sourceTableName.StartsWith(pair.Key, StringComparison.OrdinalIgnoreCase))
                    {
                        return pair.Value;
                    }
                }
            }

            return EscapeIdentifier(sourceTableName) + ".";
        }

        private static string EscapeIdentifier(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            return "[" + value.Replace("]", "]]") + "]";
        }

        private static Dictionary<string, string> GetTableQualifiers(FromClause fromClause)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (fromClause == null || fromClause.TableReferences == null)
                return result;

            foreach (TableReference tableReference in fromClause.TableReferences)
            {
                AddTableQualifiers(tableReference, result);
            }

            return result;
        }

        private static List<TableInfo> GetTables(FromClause fromClause)
        {
            var result = new List<TableInfo>();
            if (fromClause == null || fromClause.TableReferences == null)
                return result;

            foreach (TableReference tableReference in fromClause.TableReferences)
            {
                AddTables(tableReference, result);
            }

            return result;
        }

        private static void AddTables(TableReference tableReference, List<TableInfo> result)
        {
            if (tableReference == null)
                return;

            if (tableReference is NamedTableReference namedTable)
            {
                string tableName = namedTable.SchemaObject?.BaseIdentifier?.Value;
                if (string.IsNullOrWhiteSpace(tableName))
                    return;

                string alias = namedTable.Alias?.Value;
                result.Add(new TableInfo
                {
                    SchemaName = namedTable.SchemaObject?.SchemaIdentifier?.Value,
                    TableName = tableName,
                    Qualifier = !string.IsNullOrWhiteSpace(alias) ? alias + "." : EscapeIdentifier(tableName) + "."
                });
                return;
            }

            if (tableReference is QualifiedJoin qualifiedJoin)
            {
                AddTables(qualifiedJoin.FirstTableReference, result);
                AddTables(qualifiedJoin.SecondTableReference, result);
                return;
            }

            if (tableReference is JoinParenthesisTableReference parenthesizedJoin)
            {
                AddTables(parenthesizedJoin.Join, result);
            }
        }

        private static void AddTableQualifiers(TableReference tableReference, Dictionary<string, string> result)
        {
            if (tableReference == null)
                return;

            if (tableReference is NamedTableReference namedTable)
            {
                string tableName = namedTable.SchemaObject?.BaseIdentifier?.Value;
                if (string.IsNullOrWhiteSpace(tableName))
                    return;

                string alias = namedTable.Alias?.Value;
                result[tableName] = !string.IsNullOrWhiteSpace(alias) ? alias + "." : EscapeIdentifier(tableName) + ".";
                return;
            }

            if (tableReference is QualifiedJoin qualifiedJoin)
            {
                AddTableQualifiers(qualifiedJoin.FirstTableReference, result);
                AddTableQualifiers(qualifiedJoin.SecondTableReference, result);
                return;
            }

            if (tableReference is JoinParenthesisTableReference parenthesizedJoin)
            {
                AddTableQualifiers(parenthesizedJoin.Join, result);
            }
        }

        private static int GetAbsoluteOffset(string text, int targetLine, int targetColumn)
        {
            int line = 0;
            int column = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (line == targetLine && column == targetColumn)
                    return i;

                if (text[i] == '\r')
                {
                    if (i + 1 < text.Length && text[i + 1] == '\n')
                        i++;

                    line++;
                    column = 0;
                }
                else if (text[i] == '\n')
                {
                    line++;
                    column = 0;
                }
                else
                {
                    column++;
                }
            }

            return text.Length;
        }

        private static void ReplaceText(IVsTextLines textLines, int line, int startColumn, int endColumn, string replacement)
        {
            IntPtr pNewText = Marshal.StringToHGlobalUni(replacement);
            try
            {
                TextSpan[] changedSpan = new TextSpan[1];
                textLines.ReplaceLines(line, startColumn, line, endColumn, pNewText, replacement.Length, changedSpan);
            }
            finally
            {
                Marshal.FreeHGlobal(pNewText);
            }
        }

        private static void SetCaretPosition(IVsTextView textView, int startLine, int startColumn, string text, int offset)
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
    }
}
