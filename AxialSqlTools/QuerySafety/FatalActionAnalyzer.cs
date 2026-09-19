using Microsoft.SqlServer.TransactSql.ScriptDom;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AxialSqlTools.QuerySafety
{
    internal sealed class FatalAction
    {
        public string Operation { get; set; }
        public string Target { get; set; }
        public int Line { get; set; }
    }

    internal sealed class FatalActionAnalysis
    {
        public List<FatalAction> Actions { get; } = new List<FatalAction>();
        public string Limitation { get; set; }
        public bool RequiresConfirmation => Actions.Count != 0 || Limitation != null;
    }

    internal static class FatalActionAnalyzer
    {
        internal static FatalActionAnalysis Analyze(string sql)
        {
            var result = new FatalActionAnalysis();
            if (string.IsNullOrWhiteSpace(sql)) return result;

            var parser = new TSql170Parser(true);
            TSqlFragment fragment;
            IList<ParseError> errors;
            using (var reader = new StringReader(WithoutBatchRepeatCounts(sql, parser))) fragment = parser.Parse(reader, out errors);
            fragment?.Accept(new ActionVisitor(result));
            if (errors.Count != 0)
            {
                // A partially parsed script must not silently bypass the warning.
                var first = errors[0];
                result.Limitation = "The query could not be fully checked (SQL syntax or unsupported syntax near line " + first.Line +
                    "). Review it before running. Repeat approval is unavailable for this execution.";
            }
            return result;
        }

        private static string WithoutBatchRepeatCounts(string sql, TSqlParser parser)
        {
            // GO n is an SSMS batch directive. ScriptDom parses GO, but not its repeat count.
            // Use tokens so GO-like text inside strings, identifiers and comments is untouched.
            IList<ParseError> ignored;
            IList<TSqlParserToken> tokens;
            using (var reader = new StringReader(sql)) tokens = parser.GetTokenStream(reader, out ignored);
            char[] text = null;
            for (int i = 0; i < tokens.Count; i++)
            {
                if (tokens[i].TokenType != TSqlTokenType.Go) continue;
                int j = i + 1;
                while (j < tokens.Count && tokens[j].TokenType == TSqlTokenType.WhiteSpace) j++;
                if (j >= tokens.Count || tokens[j].Line != tokens[i].Line || tokens[j].TokenType != TSqlTokenType.Integer) continue;
                if (text == null) text = sql.ToCharArray();
                for (int k = 0; k < tokens[j].Text.Length; k++) text[tokens[j].Offset + k] = ' ';
            }
            return text == null ? sql : new string(text);
        }

        private sealed class ActionVisitor : TSqlFragmentVisitor
        {
            private readonly FatalActionAnalysis result;
            internal ActionVisitor(FatalActionAnalysis result) { this.result = result; }

            public override void ExplicitVisit(UpdateStatement node)
            {
                var spec = node.UpdateSpecification;
                if (spec.WhereClause == null && !IsTemporary(spec.Target, spec.FromClause, node.WithCtesAndXmlNamespaces))
                    Add("UPDATE without WHERE", spec.Target, node.StartLine);
            }

            public override void ExplicitVisit(DeleteStatement node)
            {
                var spec = node.DeleteSpecification;
                if (spec.WhereClause == null && !IsTemporary(spec.Target, spec.FromClause, node.WithCtesAndXmlNamespaces))
                    Add("DELETE without WHERE", spec.Target, node.StartLine);
            }

            public override void ExplicitVisit(TruncateTableStatement node)
            {
                if (!IsTempName(node.TableName))
                    result.Actions.Add(new FatalAction { Operation = "TRUNCATE TABLE", Target = Name(node.TableName), Line = node.StartLine });
            }

            // Defining a module does not execute its body. EXEC and dynamic SQL are not expanded.
            public override void ExplicitVisit(CreateProcedureStatement node) { }
            public override void ExplicitVisit(AlterProcedureStatement node) { }
            public override void ExplicitVisit(CreateOrAlterProcedureStatement node) { }
            public override void ExplicitVisit(CreateFunctionStatement node) { }
            public override void ExplicitVisit(AlterFunctionStatement node) { }
            public override void ExplicitVisit(CreateOrAlterFunctionStatement node) { }
            public override void ExplicitVisit(CreateTriggerStatement node) { }
            public override void ExplicitVisit(AlterTriggerStatement node) { }
            public override void ExplicitVisit(CreateOrAlterTriggerStatement node) { }

            private void Add(string operation, TableReference target, int line)
            {
                var named = target as NamedTableReference;
                result.Actions.Add(new FatalAction { Operation = operation, Target = named == null ? "Target in query" : Name(named.SchemaObject), Line = line });
            }
        }

        private static string Name(SchemaObjectName name) => name == null ? "Target in query" :
            string.Join(".", name.Identifiers.Select(i => "[" + i.Value.Replace("]", "]]") + "]"));

        private static bool IsTempName(SchemaObjectName name) => name?.BaseIdentifier?.Value.StartsWith("#", StringComparison.Ordinal) == true;

        private static bool IsTemporary(TableReference target, FromClause from, WithCtesAndXmlNamespaces ctes)
        {
            var named = target as NamedTableReference;
            if (named?.SchemaObject?.Identifiers.Count == 1 && from != null)
            {
                // Only aliases in the outer FROM can identify the DML target. Never use aliases from a subquery.
                string name = named.SchemaObject.BaseIdentifier.Value;
                var matches = Sources(from).OfType<TableReferenceWithAlias>()
                    .Where(t => string.Equals(t.Alias?.Value, name, StringComparison.OrdinalIgnoreCase)).ToList();
                if (matches.Count != 0)
                    return matches.Count == 1 && string.Equals(matches[0].Alias.Value, name, StringComparison.Ordinal) &&
                        IsTemporarySource(matches[0], ctes, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
            }
            return IsTemporarySource(target, ctes, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
        }

        private static bool IsTemporarySource(TableReference source, WithCtesAndXmlNamespaces ctes, HashSet<string> visiting)
        {
            if (source is VariableTableReference) return true;
            if (source is QueryDerivedTable derived) return IsTemporaryQuery(derived.QueryExpression, ctes, visiting);
            if (!(source is NamedTableReference named)) return false;
            if (named.SchemaObject?.Identifiers.Count == 1 && ctes != null)
            {
                string name = named.SchemaObject.BaseIdentifier.Value;
                var matches = ctes.CommonTableExpressions.Where(c => string.Equals(c.ExpressionName.Value, name, StringComparison.OrdinalIgnoreCase)).ToList();
                if (matches.Count != 0)
                {
                    // Without the database collation, a case-only match cannot prove the target is temporary.
                    if (matches.Count != 1 || matches[0].ExpressionName.Value != name || !visiting.Add(name)) return false;
                    bool temporary = IsTemporaryQuery(matches[0].QueryExpression, ctes, visiting);
                    visiting.Remove(name);
                    return temporary;
                }
            }
            return IsTempName(named.SchemaObject);
        }

        private static bool IsTemporaryQuery(QueryExpression query, WithCtesAndXmlNamespaces ctes, HashSet<string> visiting)
        {
            if (query is QueryParenthesisExpression parenthesis) return IsTemporaryQuery(parenthesis.QueryExpression, ctes, visiting);
            if (query is BinaryQueryExpression binary)
                return IsTemporaryQuery(binary.FirstQueryExpression, ctes, visiting) && IsTemporaryQuery(binary.SecondQueryExpression, ctes, visiting);
            if (!(query is QuerySpecification spec) || spec.FromClause == null) return false;
            var sources = Sources(spec.FromClause).ToList();
            return sources.Count > 0 && sources.All(s => IsTemporarySource(s, ctes, visiting));
        }

        private static IEnumerable<TableReference> Sources(FromClause from) => from.TableReferences.SelectMany(Leaves);
        private static IEnumerable<TableReference> Leaves(TableReference table)
        {
            if (table is JoinTableReference join)
                return Leaves(join.FirstTableReference).Concat(Leaves(join.SecondTableReference));
            if (table is JoinParenthesisTableReference parenthesis) return Leaves(parenthesis.Join);
            return new[] { table };
        }
    }
}
