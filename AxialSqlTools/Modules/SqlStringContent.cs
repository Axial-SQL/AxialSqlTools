using Microsoft.SqlServer.TransactSql.ScriptDom;
using System.Collections.Generic;
using System.IO;

namespace AxialSqlTools
{
    internal static class SqlStringContent
    {
        internal static bool TryFind(string sql, int position, out int start, out int length)
        {
            start = length = 0;
            if (string.IsNullOrEmpty(sql) || position < 0 || position >= sql.Length) return false;

            // ponytail: tokenize on double-click, O(document length); cache by snapshot if large scripts need it.
            IList<ParseError> errors;
            var tokens = new TSql170Parser(true).GetTokenStream(new StringReader(sql), out errors);
            foreach (var token in tokens)
            {
                if (token.Offset > position) break;
                if (token.TokenType != TSqlTokenType.AsciiStringLiteral
                    && token.TokenType != TSqlTokenType.UnicodeStringLiteral) continue;

                int contentStart = token.Offset + token.Text.IndexOf('\'') + 1;
                int contentEnd = token.Offset + token.Text.Length - 1;
                if (position < contentStart || position >= contentEnd) continue;
                start = contentStart;
                length = contentEnd - contentStart;
                return true;
            }
            return false;
        }
    }
}
