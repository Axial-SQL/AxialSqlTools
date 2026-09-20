using Microsoft.SqlServer.TransactSql.ScriptDom;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace AxialSqlTools
{
    internal static class TsqlSelectFieldIndentation
    {
        internal static string Format(string sql, TSqlParser parser, int indentSize)
        {
            IList<ParseError> errors;
            TSqlFragment fragment;
            using (var reader = new StringReader(sql)) fragment = parser.Parse(reader, out errors);
            if (errors.Count != 0) throw new InvalidOperationException("Cannot indent SELECT fields in an invalid formatted query.");

            var visitor = new SelectVisitor();
            fragment.Accept(visitor);
            var tokens = fragment.ScriptTokenStream;
            var text = tokens.Select(t => t.Text ?? "").ToArray();
            string newline = sql.Contains("\r\n") ? "\r\n" : "\n";

            // Outer lists move entire expressions first. Nested SELECTs then use their updated columns.
            foreach (var query in visitor.Queries.OrderBy(q => q.FirstTokenIndex))
            {
                int fieldIndent = ColumnBefore(text, query.FirstTokenIndex, indentSize) + indentSize;
                int gapStart = query.TopRowFilter?.LastTokenIndex + 1 ?? query.FirstTokenIndex + 1;
                if (query.TopRowFilter == null)
                    while (gapStart < query.SelectElements[0].FirstTokenIndex && tokens[gapStart - 1].TokenType != TSqlTokenType.Distinct) gapStart++;
                foreach (var field in query.SelectElements)
                {
                    int delta = fieldIndent - ColumnBefore(text, field.FirstTokenIndex, indentSize);
                    ShiftContinuationLines(tokens, text, field.FirstTokenIndex, field.LastTokenIndex, delta, indentSize);
                    AlignListComments(tokens, text, gapStart, field.FirstTokenIndex, fieldIndent, newline);
                    StartNewLine(tokens, text, field.FirstTokenIndex, fieldIndent, newline);
                    gapStart = field.LastTokenIndex + 1;
                }
            }
            return string.Concat(text);
        }

        private sealed class SelectVisitor : TSqlFragmentVisitor
        {
            internal readonly List<QuerySpecification> Queries = new List<QuerySpecification>();
            public override void ExplicitVisit(QuerySpecification node)
            {
                if (node.TopRowFilter != null || node.UniqueRowFilter == UniqueRowFilter.Distinct) Queries.Add(node);
                base.ExplicitVisit(node);
            }
        }

        private static int ColumnBefore(string[] text, int index, int indentSize)
        {
            var parts = new List<string>();
            for (int i = index - 1; i >= 0; i--)
            {
                int newline = text[i].LastIndexOf('\n');
                parts.Add(newline < 0 ? text[i] : text[i].Substring(newline + 1));
                if (newline >= 0) break;
            }
            parts.Reverse();
            return Width(string.Concat(parts), indentSize);
        }

        private static int Width(string text, int indentSize)
        {
            int column = 0;
            foreach (char c in text) column = c == '\t' ? column + indentSize - column % indentSize : column + 1;
            return column;
        }

        private static void ShiftContinuationLines(IList<TSqlParserToken> tokens, string[] text, int start, int end, int delta, int indentSize)
        {
            for (int i = start; i <= end; i++)
            {
                if (tokens[i].TokenType != TSqlTokenType.WhiteSpace) continue;
                int first = i;
                var whitespace = new StringBuilder();
                while (i <= end && tokens[i].TokenType == TSqlTokenType.WhiteSpace)
                {
                    whitespace.Append(text[i]);
                    text[i++] = "";
                }
                i--;
                string value = whitespace.ToString();
                int newline = value.LastIndexOf('\n');
                if (newline >= 0)
                {
                    int indentation = Width(value.Substring(newline + 1), indentSize);
                    value = value.Substring(0, newline + 1) + new string(' ', Math.Max(0, indentation + delta));
                }
                text[first] = value;
            }
        }

        private static void AlignListComments(IList<TSqlParserToken> tokens, string[] text, int start, int end, int indentation, string newline)
        {
            for (int i = start; i < end; i++)
            {
                if (tokens[i].TokenType != TSqlTokenType.SingleLineComment && tokens[i].TokenType != TSqlTokenType.MultilineComment) continue;
                int before = i - 1;
                bool startsOnNewLine = false;
                while (before >= start && tokens[before].TokenType == TSqlTokenType.WhiteSpace)
                {
                    startsOnNewLine |= text[before].Contains("\n");
                    before--;
                }
                // Keep inline comments in place. Standalone comments use the list's indentation,
                // without accumulating blank lines from the comment interleaver on each format pass.
                if (startsOnNewLine) StartNewLine(tokens, text, i, indentation, newline);
            }
        }

        private static void StartNewLine(IList<TSqlParserToken> tokens, string[] text, int index, int indentation, string newline)
        {
            int first = index;
            while (first > 0 && tokens[first - 1].TokenType == TSqlTokenType.WhiteSpace) first--;
            for (int i = first; i < index; i++) text[i] = "";
            // A comment can directly precede a field without a whitespace token. Keep its contents intact.
            if (first == index) text[index] = newline + new string(' ', indentation) + text[index];
            else text[first] = newline + new string(' ', indentation);
        }
    }
}
