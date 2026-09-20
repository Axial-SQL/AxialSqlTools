using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace AxialSqlTools
{
    internal static class SelectedTextSorter
    {
        internal static string Sort(string text, bool descending, bool vertical)
        {
            if (string.IsNullOrWhiteSpace(text) || text.IndexOf(',') < 0) return text;

            var items = SplitItems(text);
            if (items.Length < 2) return text;
            var numbers = items.Select(NumericKey.Parse).ToArray();
            bool numeric = numbers.All(number => number != null);
            var comparer = Comparer<int>.Create((a, b) => numeric
                ? numbers[a].CompareTo(numbers[b])
                : StringComparer.OrdinalIgnoreCase.Compare(items[a], items[b]));
            var indexes = Enumerable.Range(0, items.Length);
            var ordered = descending ? indexes.OrderByDescending(i => i, comparer) : indexes.OrderBy(i => i, comparer);
            return string.Join(vertical ? "," + Environment.NewLine : ", ", ordered.Select(i => items[i]));
        }

        private static string[] SplitItems(string text)
        {
            var items = new List<string>();
            int start = 0;
            bool inString = false;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '\'')
                {
                    // SQL escapes a quote inside a string by doubling it.
                    if (inString && i + 1 < text.Length && text[i + 1] == '\'')
                        i++;
                    else
                        inString = !inString;
                }
                else if (text[i] == ',' && !inString)
                {
                    items.Add(text.Substring(start, i - start).Trim());
                    start = i + 1;
                }
            }
            if (inString)
                throw new ArgumentException("The selection contains an unclosed string. Select complete comma-separated values, including their closing quotes.");
            items.Add(text.Substring(start).Trim());
            return items.ToArray();
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

    }
}
