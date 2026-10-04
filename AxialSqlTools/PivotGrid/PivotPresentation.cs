using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Linq;

namespace AxialSqlTools.PivotGrid
{
    internal static class PivotPresentation
    {
        public static bool RequestsEqual(PivotRequest a, PivotRequest b)
        {
            if (ReferenceEquals(a, b)) return true;
            if (a == null || b == null) return false;
            return a.NullTextIsNull == b.NullTextIsNull &&
                ArraysEqual(a.Rows, b.Rows, (x, y) => x == y) &&
                ArraysEqual(a.Columns, b.Columns, (x, y) => x == y) &&
                ArraysEqual(a.Measures, b.Measures, (x, y) => ReferenceEquals(x, y) ||
                    (x != null && y != null && x.Value == y.Value && x.Aggregation == y.Aggregation)) &&
                ArraysEqual(a.Filters, b.Filters, (x, y) => ReferenceEquals(x, y) ||
                    (x != null && y != null && x.Field == y.Field && x.Operator == y.Operator &&
                        string.Equals(x.Text, y.Text, StringComparison.Ordinal) &&
                        ArraysEqual(x.Values, y.Values, (left, right) => string.Equals(left, right, StringComparison.Ordinal))));
        }

        private static bool ArraysEqual<T>(T[] a, T[] b, Func<T, T, bool> equals)
        {
            if (ReferenceEquals(a, b)) return true;
            if (a == null || b == null || a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++)
                if (!equals(a[i], b[i])) return false;
            return true;
        }

        public static DataRowView[] SortRows(PivotResult result, PivotSnapshot source,
            PivotRequest applied, int column, ListSortDirection? direction)
        {
            if (result == null) return new DataRowView[0];

            // Return the existing views, never copies of their rows. Drill-down uses the
            // original DataRow and its index in the engine's table to identify each group.
            var table = result.Table;
            var views = table.DefaultView.Cast<DataRowView>().ToDictionary(view => view.Row);
            var rows = new List<DataRowView>(Math.Max(0, table.Rows.Count - 1));
            for (int i = 0; i < table.Rows.Count - 1; i++)
            {
                DataRowView view;
                if (views.TryGetValue(table.Rows[i], out view)) rows.Add(view);
            }

            // The final engine row is always the pinned grand total, including when a
            // source grouping label happens to be the literal text 'Grand total'.
            if (!direction.HasValue || column < 0 || column >= table.Columns.Count) return rows.ToArray();

            PivotField field = null;
            bool numeric = column >= result.RowColumnCount;
            if (!numeric && source != null && applied != null && applied.Rows != null && column < applied.Rows.Length)
            {
                int index = applied.Rows[column];
                if (index >= 0 && index < source.Fields.Length)
                {
                    field = source.Fields[index];
                    numeric = field.IsNumeric;
                }
            }

            // Compute keys once. LINQ's stable ordering preserves the original group order
            // for equal values in both directions, and resets always use the table order.
            var keyed = rows.Select(row => new KeyValuePair<DataRowView, SortKey>(row,
                CreateKey(row[column], numeric, field == null ? null : field.DataType)));
            var sorted = direction == ListSortDirection.Descending
                ? keyed.OrderByDescending(pair => pair.Value, SortKeyComparer.Instance)
                : keyed.OrderBy(pair => pair.Value, SortKeyComparer.Instance);
            return sorted.Select(pair => pair.Key).ToArray();
        }

        private static SortKey CreateKey(object value, bool numeric, Type type)
        {
            if (value == null || value == DBNull.Value) return new SortKey();
            string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
            if (numeric)
            {
                NumberKey number;
                if (NumberKey.TryParse(text, out number)) return new SortKey { Bucket = 1, Number = number };
            }
            else if (type == typeof(DateTime))
            {
                DateTime date;
                if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.RoundtripKind, out date))
                    return new SortKey { Bucket = 1, Typed = date };
            }
            else if (type == typeof(DateTimeOffset))
            {
                DateTimeOffset date;
                if (DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out date))
                    return new SortKey { Bucket = 1, Typed = date };
            }
            else if (type == typeof(TimeSpan))
            {
                TimeSpan time;
                if (TimeSpan.TryParse(text, CultureInfo.InvariantCulture, out time))
                    return new SortKey { Bucket = 1, Typed = time };
            }
            else if (type == typeof(bool))
            {
                bool boolean;
                string trimmed = text.Trim();
                if (bool.TryParse(trimmed, out boolean) || trimmed == "0" || trimmed == "1")
                    return new SortKey { Bucket = 1, Typed = trimmed == "1" || boolean };
            }

            // A separate fallback bucket makes mixed valid/invalid source labels a total,
            // transitive ordering. Pairwise numeric-or-text fallback can form cycles.
            return new SortKey { Bucket = 2, Text = text };
        }

        private sealed class SortKey
        {
            public int Bucket;
            public NumberKey Number;
            public IComparable Typed;
            public string Text;
        }

        private sealed class SortKeyComparer : IComparer<SortKey>
        {
            public static readonly SortKeyComparer Instance = new SortKeyComparer();
            public int Compare(SortKey x, SortKey y)
            {
                int bucket = x.Bucket.CompareTo(y.Bucket);
                if (bucket != 0) return bucket;
                if (x.Bucket == 0) return 0;
                if (x.Bucket == 2) return StringComparer.CurrentCultureIgnoreCase.Compare(x.Text, y.Text);
                if (x.Number != null) return x.Number.CompareTo(y.Number);
                return x.Typed.CompareTo(y.Typed);
            }
        }

        // SQL decimal grouping labels can contain 38 significant digits, while decimal
        // holds at most 29. Compare normalized digit strings to retain exact ordering
        // without rounding, overflow, or another assembly dependency. Scientific notation
        // covers floating-point labels, too. This does not change aggregation precision.
        private sealed class NumberKey : IComparable<NumberKey>
        {
            private int sign;
            private long magnitude;
            private string digits;

            public static bool TryParse(string text, out NumberKey value)
            {
                value = null;
                text = text.Trim();
                if (text.Length == 0) return false;
                int start = text[0] == '+' || text[0] == '-' ? 1 : 0;
                int exponentAt = text.IndexOfAny(new[] { 'e', 'E' }, start);
                int end = exponentAt < 0 ? text.Length : exponentAt;
                int exponent = 0;
                if (exponentAt >= 0 && !int.TryParse(text.Substring(exponentAt + 1),
                    NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out exponent)) return false;

                int point = -1;
                int firstDigit = -1;
                int digitCount = 0;
                var buffer = new char[end - start];
                for (int i = start; i < end; i++)
                {
                    char c = text[i];
                    if (c == '.' && point < 0) { point = digitCount; continue; }
                    if (c < '0' || c > '9') return false;
                    if (c != '0' && firstDigit < 0) firstDigit = digitCount;
                    buffer[digitCount++] = c;
                }
                if (digitCount == 0) return false;
                if (firstDigit < 0) { value = new NumberKey { sign = 0, digits = "" }; return true; }
                if (point < 0) point = digitCount;
                int lastDigit = digitCount - 1;
                while (buffer[lastDigit] == '0') lastDigit--;
                value = new NumberKey
                {
                    sign = text[0] == '-' ? -1 : 1,
                    magnitude = (long)point - firstDigit + exponent,
                    digits = new string(buffer, firstDigit, lastDigit - firstDigit + 1)
                };
                return true;
            }

            public int CompareTo(NumberKey other)
            {
                int comparison = sign.CompareTo(other.sign);
                if (comparison != 0 || sign == 0) return comparison;
                comparison = magnitude.CompareTo(other.magnitude);
                if (comparison == 0)
                {
                    int length = Math.Max(digits.Length, other.digits.Length);
                    for (int i = 0; i < length; i++)
                    {
                        char left = i < digits.Length ? digits[i] : '0';
                        char right = i < other.digits.Length ? other.digits[i] : '0';
                        comparison = left.CompareTo(right);
                        if (comparison != 0) break;
                    }
                }
                return sign * comparison;
            }
        }
    }
}
