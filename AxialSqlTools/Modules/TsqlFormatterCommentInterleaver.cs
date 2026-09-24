using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.SqlServer.TransactSql.ScriptDom;

public static class TsqlFormatterCommentInterleaver
{
    public static string GenerateWithComments(TSqlFragment sqlFragment, SqlScriptGenerator generator, TSqlParser parser = null)
    {
        if (sqlFragment == null) throw new ArgumentNullException(nameof(sqlFragment));
        if (generator == null) throw new ArgumentNullException(nameof(generator));
        if (parser == null) parser = new TSql170Parser(initialQuotedIdentifiers: true);
        var option = generator.Options.GetType().GetProperty("PreserveComments");
        var saved = option?.GetValue(generator.Options);
        string formatted;
        try
        {
            if (option?.CanWrite == true) option.SetValue(generator.Options, false);
            generator.GenerateScript(sqlFragment, out formatted);
        }
        finally
        {
            if (option?.CanWrite == true) option.SetValue(generator.Options, saved);
        }
        return RestoreComments(sqlFragment, formatted, parser);
    }

    internal static string RestoreComments(TSqlFragment original, string formatted, TSqlParser parser)
    {
        var target = Parse(parser, formatted);
        string restored = InterleaveComments(original, target);
        var verified = Parse(parser, restored);
        // Parsing alone is insufficient: '--' can hide an entire valid statement. Require
        // every generated code token to survive comment insertion, with its exact text.
        var expected = target.ScriptTokenStream.Where(IsCodeToken).Select(t => (t.TokenType, t.Text));
        var actual = verified.ScriptTokenStream.Where(IsCodeToken).Select(t => (t.TokenType, t.Text));
        if (!expected.SequenceEqual(actual))
            throw new InvalidOperationException("Unable to preserve comments without changing SQL tokens. The query has not been formatted.");
        var originalComments = original.ScriptTokenStream.Where(IsCommentToken).Select(t => (t.TokenType, t.Text));
        var restoredComments = verified.ScriptTokenStream.Where(IsCommentToken).Select(t => (t.TokenType, t.Text));
        if (!originalComments.SequenceEqual(restoredComments))
            throw new InvalidOperationException("Unable to preserve every comment exactly. The query has not been formatted.");
        return restored;
    }

    public static string InterleaveComments(TSqlFragment originalFragment, TSqlFragment formattedFragment)
    {
        var orig = originalFragment.ScriptTokenStream ?? throw new InvalidOperationException("Original fragment has no ScriptTokenStream.");
        var fmt = formattedFragment.ScriptTokenStream ?? throw new InvalidOperationException("Formatted fragment has no ScriptTokenStream.");
        if (fmt.Any(IsCommentToken))
            throw new InvalidOperationException("Comment insertion requires formatted SQL without comments.");
        if (!orig.Any(IsCommentToken)) return string.Concat(fmt.Select(t => t.Text));

        var originalCode = Enumerable.Range(0, orig.Count).Where(i => IsCodeToken(orig[i])).ToArray();
        var formattedCode = Enumerable.Range(0, fmt.Count).Where(i => IsCodeToken(fmt[i])).ToArray();
        var pairs = LongestCommonSubsequence(originalCode.Select(i => TokenKey(orig[i])).ToArray(),
            formattedCode.Select(i => TokenKey(fmt[i])).ToArray());
        var map = pairs.ToDictionary(pair => pair.iA, pair => pair.iB);
        var gaps = new Dictionary<int, List<Comment>>();
        int lastGap = 0;

        // Each group lives between two source code tokens. Preserve its own inline/standalone
        // layout, but never copy source indentation into the generated SQL.
        for (int code = 0; code <= originalCode.Length; code++)
        {
            int start = code == 0 ? 0 : originalCode[code - 1] + 1;
            int end = code == originalCode.Length ? orig.Count : originalCode[code];
            var comments = new List<Comment>();
            int previous = start - 1;
            for (int i = start; i < end; i++)
            {
                if (!IsCommentToken(orig[i])) continue;
                int breaks = CountNewlines(TokenText(orig, previous + 1, i));
                comments.Add(new Comment(orig[i], previous >= 0 && breaks == 0, Math.Max(0, breaks - 1)));
                previous = i;
            }
            if (comments.Count == 0) continue;

            int before = code - 1;
            while (before >= 0 && !map.ContainsKey(before)) before--;
            int after = code;
            while (after < originalCode.Length && !map.ContainsKey(after)) after++;
            int gap = comments[0].Inline && before >= 0 ? map[before] + 1
                : after < originalCode.Length ? map[after] : formattedCode.Length;
            // Put a trailing '--' after its delimiter, so an inserted semicolon or a list
            // comma cannot be swallowed by the comment or stranded on its own line.
            if (comments[0].Inline && comments[0].Token.TokenType == TSqlTokenType.SingleLineComment
                && gap < formattedCode.Length
                && (fmt[formattedCode[gap]].TokenType == TSqlTokenType.Semicolon
                    || fmt[formattedCode[gap]].TokenType == TSqlTokenType.Comma))
                gap++;
            // Adjacent source gaps may collapse when the generator removes optional tokens.
            gap = Math.Max(lastGap, gap);
            lastGap = gap;
            if (!gaps.TryGetValue(gap, out var list)) gaps[gap] = list = new List<Comment>();
            list.AddRange(comments);
        }

        string newline = fmt.Any(t => (t.Text ?? "").Contains("\r\n")) ? "\r\n" : "\n";
        var output = new StringBuilder();
        for (int gap = 0; gap <= formattedCode.Length; gap++)
        {
            int start = gap == 0 ? 0 : formattedCode[gap - 1] + 1;
            int end = gap == formattedCode.Length ? fmt.Count : formattedCode[gap];
            string whitespace = TokenText(fmt, start, end);
            bool hasNext = gap < formattedCode.Length;
            string indentation = hasNext ? IndentationBefore(fmt, formattedCode[gap], whitespace) : "";
            if (gaps.TryGetValue(gap, out var comments))
                EmitComments(output, whitespace, indentation, comments, hasNext, newline);
            else
                output.Append(whitespace);
            if (hasNext) output.Append(fmt[formattedCode[gap]].Text);
        }
        return output.ToString();
    }

    private static void EmitComments(StringBuilder output, string whitespace, string indentation,
        List<Comment> comments, bool hasNext, string newline)
    {
        bool inline = comments[0].Inline && output.Length > 0;
        for (int i = 0; i < comments.Count; i++)
        {
            var comment = comments[i];
            bool followsLineComment = i > 0 && comments[i - 1].Token.TokenType == TSqlTokenType.SingleLineComment;
            if ((i == 0 ? inline : comment.Inline) && !followsLineComment && output.Length > 0)
                output.Append(' ');
            else
            {
                int breaks = output.Length == 0 ? 0 : 1 + comment.BlankLinesBefore;
                if (i == 0) breaks = Math.Max(breaks, CountNewlines(whitespace));
                for (int n = 0; n < breaks; n++) output.Append(newline);
                output.Append(indentation);
            }
            output.Append(comment.Token.Text);
        }

        if (inline)
        {
            // The generator's gap belongs AFTER an inline comment. In particular, keep its
            // indentation for the next statement, and always terminate a '--' comment.
            if (CountNewlines(whitespace) > 0)
            {
                // Spaces before the newline would become part of the '--' token itself.
                // Keep the original comment text exact, even when a list transform emitted " \r\n".
                if (comments[comments.Count - 1].Token.TokenType == TSqlTokenType.SingleLineComment)
                    whitespace = whitespace.TrimStart(' ', '\t');
                output.Append(whitespace);
            }
            else if (hasNext && comments[comments.Count - 1].Token.TokenType == TSqlTokenType.SingleLineComment)
                output.Append(newline).Append(indentation);
            else if (hasNext) output.Append(' ');
        }
        else if (hasNext)
            output.Append(newline).Append(indentation);
    }

    private static string TokenText(IList<TSqlParserToken> tokens, int start, int end)
    {
        var text = new StringBuilder();
        for (int i = start; i < end; i++) text.Append(tokens[i].Text);
        return text.ToString();
    }

    private static string IndentationBefore(IList<TSqlParserToken> tokens, int index, string whitespace)
    {
        int newline = Math.Max(whitespace.LastIndexOf('\n'), whitespace.LastIndexOf('\r'));
        if (newline >= 0) return whitespace.Substring(newline + 1);
        // A line comment can introduce a line break where the generator used a space.
        // Align the following token to the column the generator chose for it.
        return new string(' ', Math.Max(0, tokens[index].Column - 1));
    }

    private static TSqlFragment Parse(TSqlParser parser, string sql)
    {
        using (var reader = new StringReader(sql))
        {
            var fragment = parser.Parse(reader, out var errors);
            if (errors.Count != 0)
                throw new InvalidOperationException("Comment-preserving formatting produced invalid SQL. The query has not been formatted.");
            return fragment;
        }
    }

    private static bool IsCommentToken(TSqlParserToken token) =>
        token.TokenType == TSqlTokenType.SingleLineComment || token.TokenType == TSqlTokenType.MultilineComment;

    private static bool IsCodeToken(TSqlParserToken token) =>
        !IsCommentToken(token) && token.TokenType != TSqlTokenType.WhiteSpace && token.TokenType != TSqlTokenType.EndOfFile;

    private static int CountNewlines(string text)
    {
        int count = 0;
        for (int i = 0; i < text.Length; i++)
            if (text[i] == '\n' || (text[i] == '\r' && (i + 1 == text.Length || text[i + 1] != '\n'))) count++;
        return count;
    }

    private sealed class Comment
    {
        internal readonly TSqlParserToken Token;
        internal readonly bool Inline;
        internal readonly int BlankLinesBefore;
        internal Comment(TSqlParserToken token, bool inline, int blankLinesBefore)
        {
            Token = token;
            Inline = inline;
            BlankLinesBefore = blankLinesBefore;
        }
    }

    private static (TSqlTokenType type, string norm) TokenKey(TSqlParserToken t)
    {
        // Preserve exact text for string/number literals; uppercase for everything else.
        switch (t.TokenType)
        {
            case TSqlTokenType.Execute:
                return (t.TokenType, "EXECUTE");
            case TSqlTokenType.Procedure:
                return (t.TokenType, "PROCEDURE");
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

    // Myers' bidirectional alignment uses linear working memory, unlike the old N*M
    // table. Prefix/suffix runs also make already-formatted scripts a linear-time case.
    private static List<(int iA, int iB)> LongestCommonSubsequence<T>(IList<T> a, IList<T> b)
        where T : IEquatable<T>
    {
        var result = new List<(int, int)>();
        MatchRange(a, 0, a.Count, b, 0, b.Count, result);
        return result;
    }

    private static void MatchRange<T>(IList<T> a, int aStart, int aEnd, IList<T> b, int bStart, int bEnd,
        List<(int, int)> result) where T : IEquatable<T>
    {
        while (aStart < aEnd && bStart < bEnd && a[aStart].Equals(b[bStart]))
            result.Add((aStart++, bStart++));
        int suffix = 0;
        while (aStart < aEnd && bStart < bEnd && a[aEnd - 1].Equals(b[bEnd - 1]))
        {
            aEnd--; bEnd--; suffix++;
        }
        if (aStart < aEnd && bStart < bEnd)
        {
            if (aEnd - aStart == 1)
            {
                for (int j = bStart; j < bEnd; j++)
                    if (a[aStart].Equals(b[j])) { result.Add((aStart, j)); break; }
            }
            else if (bEnd - bStart == 1)
            {
                for (int i = aStart; i < aEnd; i++)
                    if (a[i].Equals(b[bStart])) { result.Add((i, bStart)); break; }
            }
            else
            {
                var split = FindSplit(a, aStart, aEnd, b, bStart, bEnd);
                if (split.iA >= 0 && (split.iA > aStart || split.iB > bStart)
                    && (split.iA < aEnd || split.iB < bEnd))
                {
                    MatchRange(a, aStart, split.iA, b, bStart, split.iB, result);
                    MatchRange(a, split.iA, aEnd, b, split.iB, bEnd, result);
                }
            }
        }
        for (int i = 0; i < suffix; i++) result.Add((aEnd + i, bEnd + i));
    }

    private static (int iA, int iB) FindSplit<T>(IList<T> a, int aStart, int aEnd, IList<T> b, int bStart, int bEnd)
        where T : IEquatable<T>
    {
        int n = aEnd - aStart, m = bEnd - bStart;
        int maxDistance = (n + m + 1) / 2, offset = maxDistance + 1;
        var forward = Enumerable.Repeat(-1, 2 * maxDistance + 3).ToArray();
        var reverse = Enumerable.Repeat(-1, forward.Length).ToArray();
        forward[offset + 1] = reverse[offset + 1] = 0;
        int delta = n - m;
        bool odd = delta % 2 != 0;
        for (int distance = 0; distance <= maxDistance; distance++)
        {
            for (int k = -distance; k <= distance; k += 2)
            {
                int slot = offset + k;
                int x = k == -distance || (k != distance && forward[slot - 1] < forward[slot + 1])
                    ? forward[slot + 1] : forward[slot - 1] + 1;
                int y = x - k;
                while (x < n && y < m && a[aStart + x].Equals(b[bStart + y])) { x++; y++; }
                forward[slot] = x;
                int reverseK = delta - k;
                if (odd && reverseK >= -(distance - 1) && reverseK <= distance - 1
                    && reverse[offset + reverseK] >= 0 && x >= n - reverse[offset + reverseK])
                    return (aStart + x, bStart + y);
            }
            for (int k = -distance; k <= distance; k += 2)
            {
                int slot = offset + k;
                int x = k == -distance || (k != distance && reverse[slot - 1] < reverse[slot + 1])
                    ? reverse[slot + 1] : reverse[slot - 1] + 1;
                int y = x - k;
                while (x < n && y < m && a[aEnd - x - 1].Equals(b[bEnd - y - 1])) { x++; y++; }
                reverse[slot] = x;
                int forwardK = delta - k;
                if (!odd && forwardK >= -distance && forwardK <= distance
                    && forward[offset + forwardK] >= 0 && forward[offset + forwardK] >= n - x)
                {
                    int splitX = forward[offset + forwardK];
                    return (aStart + splitX, bStart + splitX - forwardK);
                }
            }
        }
        return (-1, -1);
    }
}
