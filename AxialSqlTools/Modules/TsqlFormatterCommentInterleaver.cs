using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.SqlServer.TransactSql.ScriptDom;

public static class TsqlFormatterCommentInterleaver
{
    /// <summary>
    /// Formats the original SQL with the given ScriptDom generator, then re-inserts comments
    /// from the original text into the formatted SQL using LCS alignment on non-comment tokens.
    /// </summary>
    public static string GenerateWithComments(TSqlFragment sqlFragment, SqlScriptGenerator generator, TSqlParser parser = null)
    {
        if (sqlFragment == null) throw new ArgumentNullException(nameof(sqlFragment));
        if (generator == null) throw new ArgumentNullException(nameof(generator));

        if (parser is null)
            parser = new TSql170Parser(initialQuotedIdentifiers: true);

        generator.GenerateScript(sqlFragment, out var formattedSql);

        // Re-parse formatted output
        var fmtFrag = Parse(parser, formattedSql, out _);

        return InterleaveComments(sqlFragment, fmtFrag);
    }

    /// <summary>
    /// Re-inserts comments from orig into fmt (both already parsed).
    /// </summary>
    public static string InterleaveComments(TSqlFragment originalFragment, TSqlFragment formattedFragment)
    {
        var orig = originalFragment.ScriptTokenStream ?? throw new InvalidOperationException("Original fragment has no ScriptTokenStream.");
        var fmt = formattedFragment.ScriptTokenStream ?? throw new InvalidOperationException("Formatted fragment has no ScriptTokenStream.");

        // Build lists of indices for code (non-comment, non-whitespace) tokens
        var origCodeIdx = IndicesWhere(orig, IsCodeToken);
        var fmtCodeIdx = IndicesWhere(fmt, IsCodeToken);

        // Build normalized keys for LCS (TokenType + normalized text)
        var origKeys = origCodeIdx.Select(i => TokenKey(orig[i])).ToList();
        var fmtKeys = fmtCodeIdx.Select(i => TokenKey(fmt[i])).ToList();

        // LCS between origKeys and fmtKeys
        var pairs = LongestCommonSubsequence(origKeys, fmtKeys);

        // Map: original code token index -> formatted code token index
        var mapOrigToFmt = new Dictionary<int, int>();
        foreach (var (iOrigKey, iFmtKey) in pairs)
        {
            mapOrigToFmt[origCodeIdx[iOrigKey]] = fmtCodeIdx[iFmtKey];
        }

        // Collect comments from original and schedule injection BEFORE a formatted code token.
        var injectBeforeFmtIndex = new Dictionary<int, List<CommentCluster>>();
        var eofComments = new List<CommentCluster>();

        // Helper to add a comment to a bucket
        void Schedule(int fmtIndex, CommentCluster cluster)
        {
            if (!injectBeforeFmtIndex.TryGetValue(fmtIndex, out var list))
                injectBeforeFmtIndex[fmtIndex] = list = new List<CommentCluster>();
            list.Add(cluster);
        }

        for (int i = 0; i < orig.Count; i++)
        {
            if (!IsCommentToken(orig[i])) continue;

            bool startsNewLine = StartsOnNewLine(orig, i);
            int blankLinesBefore = startsNewLine ? CountBlankLinesBefore(orig, i) : 0;

            var chunk = new StringBuilder();
            chunk.Append(orig[i].Text);
            int k = i + 1;
            while (k < orig.Count && IsWsOrComment(orig[k]))
            {
                chunk.Append(orig[k].Text);
                k++;
            }
            i = k - 1;

            // Anchor to the *next statement header* if the immediate next code token is ';'
            int anchorOrig = FindAnchorAfterComments(orig, k);

            var cluster = new CommentCluster(chunk.ToString(), startsNewLine, blankLinesBefore);

            if (anchorOrig >= 0 && mapOrigToFmt.TryGetValue(anchorOrig, out int fmtIndex))
                Schedule(fmtIndex, cluster);
            else
                eofComments.Add(cluster);
        }

        // Emit the formatted stream and inject comments at anchors (before the code token)
        var sb = new StringBuilder();
        int cur = 0;

        foreach (var jCode in fmtCodeIdx)
        {
            while (cur < jCode)
            {
                sb.Append(fmt[cur].Text);
                cur++;
            }

            if (injectBeforeFmtIndex.TryGetValue(jCode, out var toInject))
            {
                foreach (var c in toInject)
                {
                    if (c.StartsNewLine)
                    {
                        // We want: at least (1 + blankLinesBefore) newlines before the comment.
                        int required = 1 + c.BlankLinesBefore;
                        int have = TrailingNewlines(sb);
                        for (int add = have; add < required; add++)
                            sb.Append(Environment.NewLine);
                    }
                    sb.Append(c.Text);
                }
            }

            sb.Append(fmt[jCode].Text);
            cur++;
        }

        while (cur < fmt.Count) { sb.Append(fmt[cur].Text); cur++; }

        foreach (var c in eofComments)
        {
            if (c.StartsNewLine && !EndsWithNewline(sb))
                sb.Append(Environment.NewLine);
            sb.Append(c.Text);
        }

        return sb.ToString();
    }

    // ----------------- Helpers -----------------

    private static TSqlFragment Parse(TSqlParser parser, string sql, out IList<ParseError> errors)
    {
        using (var sr = new StringReader(sql))
        {
            var frag = parser.Parse(sr, out errors);
            return frag;
        }
    }

    private static List<int> IndicesWhere(IList<TSqlParserToken> tokens, Func<TSqlParserToken, bool> pred)
    {
        var list = new List<int>(tokens.Count);
        for (int i = 0; i < tokens.Count; i++)
            if (pred(tokens[i])) list.Add(i);
        return list;
    }

    private static int NextIndex(IList<TSqlParserToken> tokens, int start, Func<TSqlParserToken, bool> pred)
    {
        for (int i = start; i < tokens.Count; i++)
            if (pred(tokens[i])) return i;
        return -1;
    }

    private static bool IsCommentToken(TSqlParserToken t) =>
        t.TokenType == TSqlTokenType.SingleLineComment ||
        t.TokenType == TSqlTokenType.MultilineComment;

    private static bool IsWhitespaceToken(TSqlParserToken t) =>
        t.TokenType == TSqlTokenType.WhiteSpace;

    private static bool IsCodeToken(TSqlParserToken t) =>
        !IsCommentToken(t) && !IsWhitespaceToken(t);

    private static bool HasNewline(string s) =>
        s.IndexOf('\n') >= 0 || s.IndexOf('\r') >= 0;

    private static bool StartsOnNewLine(IList<TSqlParserToken> tokens, int commentIndex)
    {
        // If the comment is the first token, it's at line start.
        if (commentIndex <= 0) return true;

        // Walk left across whitespace; if any contains a newline, the comment starts on a new line.
        int k = commentIndex - 1;
        bool sawWs = false;
        while (k >= 0 && IsWhitespaceToken(tokens[k]))
        {
            sawWs = true;
            if (HasNewline(tokens[k].Text)) return true;
            k--;
        }

        // If there was no whitespace, it's immediately after another token on the same line.
        // If there was whitespace but no newline, also same line.
        return false;
    }

    private static bool EndsWithNewline(StringBuilder sb)
    {
        if (sb.Length == 0) return false;
        char last = sb[sb.Length - 1];
        return last == '\n' || last == '\r';
    }

    private static bool IsWsOrComment(TSqlParserToken t) =>
        t.TokenType == TSqlTokenType.WhiteSpace ||
        t.TokenType == TSqlTokenType.SingleLineComment ||
        t.TokenType == TSqlTokenType.MultilineComment;


    // Normalized comparison key for LCS: (TokenType, normalized text).
    // For identifiers/keywords, T-SQL is generally case-insensitive; normalize to upper.
    private static (TSqlTokenType type, string norm) TokenKey(TSqlParserToken t)
    {
        // Preserve exact text for string/number literals; uppercase for everything else.
        switch (t.TokenType)
        {
            case TSqlTokenType.AsciiStringLiteral:
            case TSqlTokenType.UnicodeStringLiteral:
            case TSqlTokenType.Integer:
            case TSqlTokenType.Real:
            case TSqlTokenType.HexLiteral:
                return (t.TokenType, t.Text); // keep exact
            default:
                if (!string.IsNullOrEmpty(t.Text))
                    return (t.TokenType, t.Text.ToUpperInvariant());
                else
                    return (t.TokenType, t.Text);
        }
    }

    private static bool IsSemicolonToken(TSqlParserToken t) =>
        t.TokenType == TSqlTokenType.Semicolon || t.Text == ";";

    private static bool IsStmtHeaderToken(TSqlParserToken t)
    {
        // Use text so it works across ScriptDom versions.
        var k = t.Text.ToUpperInvariant();
        switch (k)
        {
            case "WITH":
            case "SELECT":
            case "INSERT":
            case "UPDATE":
            case "DELETE":
            case "MERGE":
            case "CREATE":
            case "ALTER":
            case "DROP":
            case "EXEC":
            case "EXECUTE":
            case "DECLARE":
            case "BEGIN":
            case "IF":
            case "WHILE":
            case "RETURN":
            case "TRUNCATE":
            case "USE":
                return true;
            default:
                return false;
        }
    }

    private static int CountNewlines(string s)
    {
        int c = 0;
        for (int i = 0; i < s.Length; i++)
            if (s[i] == '\n') c++;
        return c;
    }

    private static int CountBlankLinesBefore(IList<TSqlParserToken> tokens, int commentIndex)
    {
        // Count newlines in whitespace between the previous CODE token and the comment.
        int prevCode = -1;
        for (int i = commentIndex - 1; i >= 0; i--)
        {
            if (IsCodeToken(tokens[i])) { prevCode = i; break; }
            if (!IsWhitespaceToken(tokens[i]) && !IsCommentToken(tokens[i])) break;
        }

        int newlines = 0;
        for (int i = prevCode + 1; i < commentIndex; i++)
            if (IsWhitespaceToken(tokens[i])) newlines += CountNewlines(tokens[i].Text);

        // One newline ends the previous line; extras are blank lines.
        return Math.Max(0, newlines - 1);
    }

    private static int FindAnchorAfterComments(IList<TSqlParserToken> tokens, int start)
    {
        // First non-comment, non-whitespace after the cluster
        int idx = NextIndex(tokens, start, IsCodeToken);
        if (idx < 0) return -1;

        // If it's a semicolon, prefer the *next* statement header if available.
        if (IsSemicolonToken(tokens[idx]))
        {
            int next = NextIndex(tokens, idx + 1, IsCodeToken);
            if (next >= 0 && IsStmtHeaderToken(tokens[next]))
                return next; // anchor to WITH / SELECT / etc.
        }

        return idx;
    }

    private static int TrailingNewlines(StringBuilder sb)
    {
        int c = 0;
        for (int i = sb.Length - 1; i >= 0; i--)
        {
            char ch = sb[i];
            if (ch == '\n') c++;
            else if (ch == '\r') continue;
            else break;
        }
        return c;
    }

    private sealed class CommentCluster
    {
        public string Text { get; }
        public bool StartsNewLine { get; }
        public int BlankLinesBefore { get; }
        public CommentCluster(string text, bool startsNewLine, int blankLinesBefore)
        {
            Text = text;
            StartsNewLine = startsNewLine;
            BlankLinesBefore = blankLinesBefore;
        }
    }

    /// <summary>
    /// Classic LCS over two sequences of keys. Returns list of matched index pairs (iA, iB) in order.
    /// </summary>
    private static List<(int iA, int iB)> LongestCommonSubsequence<T>(IList<T> a, IList<T> b)
        where T : IEquatable<T>
    {
        int n = a.Count, m = b.Count;
        var dp = new int[n + 1, m + 1];

        for (int i = n - 1; i >= 0; i--)
            for (int j = m - 1; j >= 0; j--)
                dp[i, j] = a[i].Equals(b[j]) ? dp[i + 1, j + 1] + 1 : Math.Max(dp[i + 1, j], dp[i, j + 1]);

        var result = new List<(int, int)>();
        int x = 0, y = 0;
        while (x < n && y < m)
        {
            if (a[x].Equals(b[y]))
            {
                result.Add((x, y));
                x++; y++;
            }
            else if (dp[x + 1, y] >= dp[x, y + 1])
                x++;
            else
                y++;
        }
        return result;
    }
}