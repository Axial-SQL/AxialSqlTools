using Microsoft.SqlServer.TransactSql.ScriptDom;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AxialSqlTools
{
    internal sealed class SqlObjectName
    {
        private SqlObjectName(string databaseName, string schemaName, string objectName)
        {
            DatabaseName = databaseName;
            SchemaName = schemaName;
            ObjectName = objectName;
        }

        public string DatabaseName { get; }
        public string SchemaName { get; }
        public string ObjectName { get; }

        public static SqlObjectName Parse(string text)
        {
            var tokens = Tokenize(text ?? string.Empty, true)
                .Where(t => t.TokenType != TSqlTokenType.WhiteSpace && t.TokenType != TSqlTokenType.EndOfFile)
                .ToList();
            if (tokens.Count > 0 && tokens[tokens.Count - 1].TokenType == TSqlTokenType.Semicolon)
                tokens.RemoveAt(tokens.Count - 1);

            var parts = new List<string>();
            string part = null;
            foreach (var token in tokens)
            {
                if (IsIdentifier(token) && part == null)
                    part = Decode(token.Text);
                else if (token.TokenType == TSqlTokenType.Dot)
                {
                    parts.Add(part);
                    part = null;
                }
                else
                    throw InvalidName();
            }
            parts.Add(part);

            if (parts.Count > 3)
                throw new ArgumentException("Use an object, schema.object, or database.schema.object name. Linked-server names are not supported.");
            if (string.IsNullOrEmpty(parts[0]) || string.IsNullOrEmpty(part)
                || parts.Any(p => p != null && (p.Length == 0 || p.Length > 128)))
                throw InvalidName();

            if (part.StartsWith("#", StringComparison.Ordinal) || part.StartsWith("@", StringComparison.Ordinal))
                throw new ArgumentException("Temp tables and table variables cannot be located in Object Explorer. Select a permanent database object.");

            return new SqlObjectName(parts.Count == 3 ? parts[0] : null,
                parts.Count >= 2 ? parts[parts.Count - 2] : null, part);
        }

        // Tokenize the document, not just its current line, so a caret inside a multi-line
        // comment or string is never treated as an object reference. No valid AST is needed.
        public static string GetNameAtCaret(string sql, int offset)
        {
            if (string.IsNullOrEmpty(sql) || offset < 0 || offset > sql.Length)
                return null;
            var tokens = Tokenize(sql);
            var atCaret = tokens.FirstOrDefault(t => t.Offset <= offset && offset < t.Offset + (t.Text?.Length ?? 0));
            if (atCaret != null && (atCaret.TokenType == TSqlTokenType.SingleLineComment
                || atCaret.TokenType == TSqlTokenType.MultilineComment
                || atCaret.TokenType == TSqlTokenType.AsciiStringLiteral
                || atCaret.TokenType == TSqlTokenType.UnicodeStringLiteral))
                return null;

            bool quotedIdentifiers = true;
            for (int i = 0; i < tokens.Count; i++)
            {
                if (tokens[i].TokenType == TSqlTokenType.Set)
                {
                    int setting = NextToken(tokens, i + 1);
                    int value = NextToken(tokens, setting + 1);
                    if (value < tokens.Count && string.Equals(tokens[setting].Text, "QUOTED_IDENTIFIER", StringComparison.OrdinalIgnoreCase))
                    {
                        if (tokens[value].TokenType == TSqlTokenType.Off) quotedIdentifiers = false;
                        else if (tokens[value].TokenType == TSqlTokenType.On) quotedIdentifiers = true;
                    }
                }
                if (!IsIdentifier(tokens[i], quotedIdentifiers))
                    continue;
                int last = i;
                int next = NextToken(tokens, i + 1);
                while (next < tokens.Count && tokens[next].TokenType == TSqlTokenType.Dot)
                {
                    last = next;
                    next = NextToken(tokens, next + 1);
                    if (next < tokens.Count && IsIdentifier(tokens[next], quotedIdentifiers))
                    {
                        last = next;
                        next = NextToken(tokens, next + 1);
                    }
                }

                int start = tokens[i].Offset;
                int end = tokens[last].Offset + tokens[last].Text.Length;
                if (offset >= start && offset <= end)
                    return sql.Substring(start, end - start);
                i = last;
            }
            return null;
        }

        internal static string Quote(string name) => "[" + name.Replace("]", "]]") + "]";

        private static IList<TSqlParserToken> Tokenize(string text, bool strict = false)
        {
            IList<ParseError> errors;
            var tokens = new TSql170Parser(true).GetTokenStream(new StringReader(text), out errors);
            if (strict && errors.Count > 0) throw InvalidName();
            return tokens;
        }

        private static bool IsIdentifier(TSqlParserToken token, bool quotedIdentifiers = true) =>
            token.TokenType == TSqlTokenType.Identifier || token.TokenType == TSqlTokenType.QuotedIdentifier
            || (quotedIdentifiers && token.TokenType == TSqlTokenType.AsciiStringOrQuotedIdentifier);

        private static int NextToken(IList<TSqlParserToken> tokens, int index)
        {
            while (index < tokens.Count && tokens[index].TokenType == TSqlTokenType.WhiteSpace)
                index++;
            return index;
        }

        private static string Decode(string text)
        {
            if (text.StartsWith("[", StringComparison.Ordinal))
                return text.Substring(1, text.Length - 2).Replace("]]", "]");
            if (text.StartsWith("\"", StringComparison.Ordinal))
                return text.Substring(1, text.Length - 2).Replace("\"\"", "\"");
            return text;
        }

        private static ArgumentException InvalidName() => new ArgumentException(
            "Select one object name, such as dbo.Orders or [Sales DB].[dbo].[Order Details].");
    }
}
