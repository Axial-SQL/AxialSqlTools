using Microsoft.SqlServer.TransactSql.ScriptDom;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace AxialSqlTools
{
    internal static class SelectedTextSorter
    {
        internal static string Sort(string text, bool descending)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;

            IList<ParseError> errors;
            var tokens = new TSql170Parser(true).GetTokenStream(new StringReader(text), out errors);
            if (errors.Count != 0)
                throw new ArgumentException("Select complete values or expressions, including their quotes and brackets.");

            var commas = new List<int>();
            var newlines = new List<int>();
            int depth = 0;
            foreach (var token in tokens)
            {
                if (token.TokenType == TSqlTokenType.SingleLineComment || token.TokenType == TSqlTokenType.MultilineComment)
                    throw new ArgumentException("Select the list without SQL comments so comments cannot move to a different value.");
                if (token.TokenType == TSqlTokenType.LeftParenthesis) depth++;
                if (token.TokenType == TSqlTokenType.RightParenthesis && --depth < 0)
                    throw new ArgumentException("Select complete expressions without the surrounding IN clause parentheses.");
                if (depth != 0) continue;
                if (token.TokenType == TSqlTokenType.Semicolon)
                    throw new ArgumentException("Select only the fields or values to sort, without the statement terminator.");
                if (token.TokenType == TSqlTokenType.Comma) commas.Add(token.Offset);
                if (token.TokenType == TSqlTokenType.WhiteSpace)
                {
                    for (int i = 0; i < token.Text.Length; i++)
                    {
                        if (token.Text[i] == '\r' || token.Text[i] == '\n')
                        {
                            newlines.Add(token.Offset + i);
                            if (token.Text[i] == '\r' && i + 1 < token.Text.Length && token.Text[i + 1] == '\n') i++;
                        }
                    }
                }
            }
            if (depth != 0)
                throw new ArgumentException("Select complete expressions with matching parentheses.");

            // Only top-level commas delimit a SQL list. With no commas, sort lines.
            var separators = commas.Count > 0 ? commas : newlines;
            var slots = new List<Slot>();
            int start = 0;
            foreach (int end in separators.Concat(new[] { text.Length }))
            {
                int first = start, last = end;
                while (first < last && char.IsWhiteSpace(text[first])) first++;
                while (last > first && char.IsWhiteSpace(text[last - 1])) last--;
                if (first < last)
                    slots.Add(new Slot { Start = first, End = last, Value = text.Substring(first, last - first) });
                else if (commas.Count > 0 && start != 0 && end != text.Length)
                    throw new ArgumentException("The selected list contains an empty value between commas.");
                start = end + 1;
            }
            if (slots.Count < 2) return text;

            // Use a single comparison mode for the whole list to keep ordering transitive.
            var numbers = slots.Select(s => NumericKey.Parse(s.Value)).ToArray();
            bool numeric = numbers.All(n => n != null);
            var keys = slots.Select(s => GetTextKey(s.Value)).ToArray();
            var comparer = Comparer<int>.Create((a, b) => numeric
                ? numbers[a].CompareTo(numbers[b])
                : StringComparer.OrdinalIgnoreCase.Compare(keys[a], keys[b]));
            var indexes = Enumerable.Range(0, slots.Count);
            var ordered = (descending ? indexes.OrderByDescending(i => i, comparer) : indexes.OrderBy(i => i, comparer)).ToArray();

            // Whitespace and delimiters stay in their original positions; only items move.
            var result = new StringBuilder(text.Length);
            int position = 0;
            for (int i = 0; i < slots.Count; i++)
            {
                result.Append(text, position, slots[i].Start - position);
                result.Append(slots[ordered[i]].Value);
                position = slots[i].End;
            }
            result.Append(text, position, text.Length - position);
            return result.ToString();
        }

        private static string GetTextKey(string value)
        {
            IList<ParseError> errors;
            var tokens = new TSql170Parser(true).GetTokenStream(new StringReader(value), out errors)
                .Where(t => t.TokenType != TSqlTokenType.WhiteSpace && t.TokenType != TSqlTokenType.EndOfFile).ToArray();
            if (tokens.Length != 1) return value;
            if (tokens[0].TokenType == TSqlTokenType.UnicodeStringLiteral) value = value.Substring(1);
            if (value.Length >= 2)
            {
                if (value[0] == '\'' && value[value.Length - 1] == '\'')
                    return value.Substring(1, value.Length - 2).Replace("''", "'");
                if (value[0] == '[' && value[value.Length - 1] == ']')
                    return value.Substring(1, value.Length - 2).Replace("]]", "]");
                if (value[0] == '"' && value[value.Length - 1] == '"')
                    return value.Substring(1, value.Length - 2).Replace("\"\"", "\"");
            }
            return value;
        }

        // Compare decimal/scientific literals exactly, including SQL decimal(38, s),
        // without floating-point rounding or dependence on the Windows locale.
        private sealed class NumericKey : IComparable<NumericKey>
        {
            private static readonly Regex Pattern = new Regex(@"^([+-]?)([0-9]*)(?:\.([0-9]*))?(?:[eE]([+-]?[0-9]+))?$", RegexOptions.CultureInvariant);
            private int sign;
            private long magnitude;
            private string digits;

            internal static NumericKey Parse(string value)
            {
                var match = Pattern.Match(value);
                if (!match.Success || match.Groups[2].Length + match.Groups[3].Length == 0) return null;
                int exponent = 0;
                if (match.Groups[4].Success && !int.TryParse(match.Groups[4].Value, NumberStyles.AllowLeadingSign,
                    CultureInfo.InvariantCulture, out exponent)) return null;
                string digits = (match.Groups[2].Value + match.Groups[3].Value).TrimStart('0');
                return new NumericKey
                {
                    sign = digits.Length == 0 ? 0 : match.Groups[1].Value == "-" ? -1 : 1,
                    magnitude = (long)exponent - match.Groups[3].Length + digits.Length,
                    digits = digits
                };
            }

            public int CompareTo(NumericKey other)
            {
                int comparison = sign.CompareTo(other.sign);
                if (comparison != 0 || sign == 0) return comparison;
                comparison = magnitude.CompareTo(other.magnitude);
                if (comparison != 0) return sign * comparison;
                for (int i = 0; i < Math.Max(digits.Length, other.digits.Length); i++)
                {
                    char left = i < digits.Length ? digits[i] : '0';
                    char right = i < other.digits.Length ? other.digits[i] : '0';
                    comparison = left.CompareTo(right);
                    if (comparison != 0) return sign * comparison;
                }
                return 0;
            }
        }

        private sealed class Slot
        {
            internal int Start;
            internal int End;
            internal string Value;
        }
    }
}
