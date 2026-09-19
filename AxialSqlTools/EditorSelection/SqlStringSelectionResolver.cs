using Microsoft.SqlServer.TransactSql.ScriptDom;
using System;
using System.Collections.Generic;
using System.IO;

namespace AxialSqlTools.EditorSelection
{
    internal struct SqlStringContentSpan
    {
        internal int Start { get; }
        internal int Length { get; }
        internal int End => Start + Length;

        internal SqlStringContentSpan(int start, int length)
        {
            Start = start;
            Length = length;
        }
    }

    // Independent of the editor so SQL boundaries can be tested without SSMS.
    internal sealed class SqlStringSelectionResolver
    {
        private readonly List<SqlStringContentSpan> _spans;

        private SqlStringSelectionResolver(List<SqlStringContentSpan> spans)
        {
            _spans = spans;
        }

        internal static SqlStringSelectionResolver Create(string sql)
        {
            var spans = new List<SqlStringContentSpan>();
            if (string.IsNullOrEmpty(sql)) return new SqlStringSelectionResolver(spans);

            IList<ParseError> errors;
            IList<TSqlParserToken> tokens;
            using (var reader = new StringReader(sql))
                tokens = new TSql170Parser(true).GetTokenStream(reader, out errors);

            foreach (var token in tokens)
            {
                if (token.TokenType != TSqlTokenType.AsciiStringLiteral
                    && token.TokenType != TSqlTokenType.UnicodeStringLiteral)
                    continue;

                string text = token.Text;
                int prefixLength = token.TokenType == TSqlTokenType.UnicodeStringLiteral ? 2 : 1;
                if (string.IsNullOrEmpty(text) || text.Length <= prefixLength + 1
                    || text[prefixLength - 1] != '\'' || text[text.Length - 1] != '\'')
                    continue;

                // A token ending in an escaped quote can still be an unfinished literal.
                bool complete = false;
                for (int i = prefixLength; i < text.Length; i++)
                {
                    if (text[i] != '\'') continue;
                    if (i + 1 < text.Length && text[i + 1] == '\'') { i++; continue; }
                    complete = i == text.Length - 1;
                    break;
                }
                if (!complete) continue;

                bool hasLexicalError = false;
                foreach (var error in errors)
                {
                    if (error.Offset >= token.Offset && error.Offset < token.Offset + text.Length)
                    {
                        hasLexicalError = true;
                        break;
                    }
                }
                if (hasLexicalError) continue;

                // Keep the original source text, including doubled quotes and newlines.
                spans.Add(new SqlStringContentSpan(token.Offset + prefixLength,
                    text.Length - prefixLength - 1));
            }
            return new SqlStringSelectionResolver(spans);
        }

        internal bool TryGetContentSpan(int position, out SqlStringContentSpan span)
        {
            int low = 0;
            int high = _spans.Count - 1;
            while (low <= high)
            {
                int middle = low + (high - low) / 2;
                var candidate = _spans[middle];
                if (position < candidate.Start) high = middle - 1;
                else if (position >= candidate.End) low = middle + 1;
                else
                {
                    span = candidate;
                    return true;
                }
            }
            span = default(SqlStringContentSpan);
            return false;
        }
    }
}
