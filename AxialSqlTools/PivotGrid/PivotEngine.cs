using NReco.PivotData;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace AxialSqlTools.PivotGrid
{
    internal sealed class PivotField
    {
        public int Index { get; }
        public string Key => "f" + Index.ToString(CultureInfo.InvariantCulture);
        public string Name { get; }
        public Type DataType { get; }
        public string Label { get; }
        public bool IsNumeric { get; }

        public PivotField(int index, string name, Type type)
        {
            Index = index;
            Name = name ?? "";
            DataType = type ?? typeof(string);
            // Ordinals keep duplicate and unnamed SQL columns unambiguous in the selectors.
            Label = (index < 0 ? "" : (index + 1) + ": ") + (string.IsNullOrEmpty(name) ? "(No column name)" : name);
            IsNumeric = type == typeof(byte) || type == typeof(short) || type == typeof(int) ||
                type == typeof(long) || type == typeof(decimal) || type == typeof(float) || type == typeof(double);
        }
    }

    internal sealed class PivotSnapshot
    {
        public PivotField[] Fields { get; }
        public string[][] Rows { get; }
        public PivotSnapshot(PivotField[] fields, string[][] rows) { Fields = fields; Rows = rows; }
    }

    internal enum PivotAggregation { CountRows, CountValues, DistinctCount, Sum, Average, Minimum, Maximum }

    internal sealed class PivotMeasure
    {
        public int Value { get; set; } = -1;
        public PivotAggregation Aggregation { get; set; }

        public PivotMeasure Copy() => new PivotMeasure { Value = Value, Aggregation = Aggregation };

        public string GetLabel(PivotSnapshot source)
        {
            if (Aggregation == PivotAggregation.CountRows) return "Count rows";
            var field = source.Fields[Value];
            string name = string.IsNullOrEmpty(field.Name) ? field.Label : field.Name;
            if (source.Fields.Count(f => string.Equals(f.Name, field.Name, StringComparison.OrdinalIgnoreCase)) > 1)
                name = field.Label;
            string operation;
            switch (Aggregation)
            {
                case PivotAggregation.CountValues: operation = "Count values"; break;
                case PivotAggregation.DistinctCount: operation = "Distinct count"; break;
                case PivotAggregation.Sum: operation = "Sum"; break;
                case PivotAggregation.Average: operation = "Average"; break;
                case PivotAggregation.Minimum: operation = "Minimum"; break;
                case PivotAggregation.Maximum: operation = "Maximum"; break;
                default: throw new InvalidOperationException("Choose a valid aggregation.");
            }
            return operation + "(" + name + ")";
        }
    }

    internal enum PivotFilterOperator
    {
        Contains, Equals, NotEquals, In, IsNull, IsNotNull,
        GreaterThan, GreaterThanOrEqual, LessThan, LessThanOrEqual
    }

    internal sealed class PivotFilter
    {
        public int Field { get; set; }
        public PivotFilterOperator Operator { get; set; }
        public string Text { get; set; } = "";
        public string[] Values { get; set; } = new string[0];

        public PivotFilter Copy() => new PivotFilter
        {
            Field = Field, Operator = Operator, Text = Text,
            Values = Values == null ? null : (string[])Values.Clone()
        };
    }

    internal sealed class PivotRequest
    {
        public int[] Rows { get; set; } = new int[0];
        public int[] Columns { get; set; } = new int[0];
        public PivotMeasure[] Measures { get; set; } = new[] { new PivotMeasure() };
        public bool NullTextIsNull { get; set; } = true;
        public PivotFilter[] Filters { get; set; } = new PivotFilter[0];

        internal PivotRequest Copy() => new PivotRequest
        {
            Rows = Rows == null ? null : (int[])Rows.Clone(),
            Columns = Columns == null ? null : (int[])Columns.Clone(),
            Measures = Measures?.Select(m => m?.Copy()).ToArray(),
            NullTextIsNull = NullTextIsNull,
            Filters = Filters?.Select(f => f?.Copy()).ToArray()
        };
    }

    internal sealed class PivotResult
    {
        public DataTable Table { get; }
        public int MatchedRows { get; }
        public int RowColumnCount => Math.Max(1, request.Rows.Length);
        private readonly PivotSnapshot source;
        private readonly PivotRequest request;
        private readonly object[][] rowKeys, columnKeys;
        private readonly Func<string[], bool> matchesFilter;

        public PivotResult(DataTable table, int matchedRows, PivotSnapshot source, PivotRequest request,
            ValueKey[] rowKeys, ValueKey[] columnKeys)
        {
            Table = table; MatchedRows = matchedRows; this.source = source; this.request = request;
            this.rowKeys = rowKeys.Select(k => (object[])k.DimKeys.Clone()).ToArray();
            this.columnKeys = columnKeys.Select(k => (object[])k.DimKeys.Clone()).ToArray();
            matchesFilter = PivotEngine.CreateFilterMatcher(source, request);
        }

        public bool CanDrillDown(int row, int column) => row >= 0 && row < Table.Rows.Count &&
            column >= RowColumnCount && column < Table.Columns.Count;

        public bool IsTotalColumn(int column) => request.Columns.Length > 0 &&
            column >= RowColumnCount + columnKeys.Length * request.Measures.Length && column < Table.Columns.Count;

        public PivotDetails GetUnderlyingRows(int row, int column, CancellationToken cancellation)
        {
            cancellation.ThrowIfCancellationRequested();
            if (!CanDrillDown(row, column)) throw new ArgumentOutOfRangeException(nameof(column), "Select a value or total cell.");
            var rowKey = row < rowKeys.Length ? rowKeys[row] : null;
            int pivotColumn = (column - RowColumnCount) / request.Measures.Length;
            var columnKey = pivotColumn < columnKeys.Length ? columnKeys[pivotColumn] : null;
            var indexes = new List<int>();
            bool Matches(string[] record, int[] fields, object[] key)
            {
                if (key == null) return true; // Totals remove the corresponding axis restriction.
                for (int i = 0; i < fields.Length; i++)
                    if (!Equals(PivotEngine.ReadValue(source, request, record, fields[i]), key[i])) return false;
                return true;
            }
            for (int i = 0; i < source.Rows.Length; i++)
            {
                cancellation.ThrowIfCancellationRequested();
                var record = source.Rows[i];
                if (matchesFilter(record) && Matches(record, request.Rows, rowKey) &&
                    Matches(record, request.Columns, columnKey)) indexes.Add(i);
            }
            var labels = new List<string>();
            void Describe(int[] fields, object[] key)
            {
                if (key == null) return;
                for (int i = 0; i < fields.Length; i++)
                    labels.Add(source.Fields[fields[i]].Label + " = " + PivotEngine.Label(key[i]));
            }
            Describe(request.Rows, rowKey);
            Describe(request.Columns, columnKey);
            if (labels.Count == 0) labels.Add("All groups");
            foreach (var filter in request.Filters)
                labels.Add("Filter: " + PivotEngine.DescribeFilter(source, filter));
            labels.Add(request.Measures[(column - RowColumnCount) % request.Measures.Length].GetLabel(source));
            return new PivotDetails(source, indexes.ToArray(), string.Join(" | ", labels));
        }
    }

    internal sealed class PivotDetails
    {
        private readonly PivotSnapshot source;
        private readonly int[] indexes;
        public string Description { get; }
        public int Count => indexes.Length;
        public int PageSize => Math.Max(1, Math.Min(200, 50000 / (source.Fields.Length + 1)));
        public int PageCount => Math.Max(1, (Count + PageSize - 1) / PageSize);
        public PivotDetails(PivotSnapshot source, int[] indexes, string description)
        { this.source = source; this.indexes = indexes; Description = description; }

        public DataTable GetPage(int page)
        {
            if (page < 0 || page >= PageCount) throw new ArgumentOutOfRangeException(nameof(page));
            var table = new DataTable { Locale = CultureInfo.InvariantCulture };
            table.Columns.Add(new DataColumn("sourceRow", typeof(int)) { Caption = "Source row" });
            foreach (var field in source.Fields)
                table.Columns.Add(new DataColumn(field.Key, typeof(string)) { Caption = field.Label });
            int end = Math.Min(Count, (page + 1) * PageSize);
            for (int i = page * PageSize; i < end; i++)
            {
                var record = table.NewRow();
                record[0] = indexes[i] + 1;
                for (int field = 0; field < source.Fields.Length; field++)
                    record[field + 1] = (object)source.Rows[indexes[i]][field] ?? DBNull.Value;
                table.Rows.Add(record);
            }
            return table;
        }
    }

    internal static class PivotEngine
    {
        internal const int MaxColumns = 200;
        internal const int MaxCells = 500000;
        internal const int MaxGroups = 100000;

        internal static bool RequiresNumber(PivotAggregation aggregation) => aggregation >= PivotAggregation.Sum;

        public static void Validate(PivotSnapshot source, PivotRequest request)
        {
            if (source == null || source.Fields == null || source.Rows == null || source.Fields.Length == 0)
                throw new InvalidOperationException("Capture a result grid before building a pivot.");
            if (request == null || request.Rows == null || request.Columns == null ||
                request.Measures == null || request.Filters == null)
                throw new InvalidOperationException("The pivot layout is incomplete. Reset the layout and try again.");
            var dimensions = request.Rows.Concat(request.Columns).ToArray();
            if (dimensions.Distinct().Count() != dimensions.Length)
                throw new InvalidOperationException("Choose different fields for Rows and Columns.");
            if (dimensions.Length > 6)
                throw new InvalidOperationException("Choose up to six grouping fields in total.");
            if (dimensions.Any(i => i < 0 || i >= source.Fields.Length))
                throw new InvalidOperationException("Select valid grid fields.");
            if (request.Measures.Length == 0)
                throw new InvalidOperationException("Add at least one value, such as Count rows.");
            if (request.Measures.Length + Math.Max(1, request.Rows.Length) > MaxColumns)
                throw new InvalidOperationException("Choose fewer values. A pivot can display at most 200 columns.");
            foreach (var measure in request.Measures)
            {
                if (measure == null || !Enum.IsDefined(typeof(PivotAggregation), measure.Aggregation))
                    throw new InvalidOperationException("Choose a valid aggregation for every value.");
                if (measure.Aggregation == PivotAggregation.CountRows) continue;
                if (measure.Value < 0 || measure.Value >= source.Fields.Length)
                    throw new InvalidOperationException("Choose a field for every value.");
                if (RequiresNumber(measure.Aggregation) && !source.Fields[measure.Value].IsNumeric)
                    throw new InvalidOperationException(measure.GetLabel(source) + " requires a numeric field.");
            }
            // Parse criteria before any source rows are read, including when the result is empty.
            CreateFilterMatcher(source, request);
        }

        public static PivotResult Build(PivotSnapshot source, PivotRequest request, CancellationToken cancellation)
        {
            cancellation.ThrowIfCancellationRequested();
            // Keep calculation and drill-down tied to a deep copy of the applied configuration.
            request = request?.Copy();
            Validate(source, request);
            var dimensions = request.Rows.Concat(request.Columns).ToArray();
            var matchesFilter = CreateFilterMatcher(source, request);
            var matchingIndexes = new List<int>();
            var rowKeys = new HashSet<ValueKey>();
            var columnKeys = new HashSet<ValueKey>();
            var groupKeys = new HashSet<ValueKey>();
            for (int i = 0; i < source.Rows.Length; i++)
            {
                cancellation.ThrowIfCancellationRequested();
                var row = source.Rows[i];
                if (!matchesFilter(row)) continue;
                rowKeys.Add(new ValueKey(request.Rows.Select(f => ReadValue(source, request, row, f)).ToArray()));
                columnKeys.Add(new ValueKey(request.Columns.Select(f => ReadValue(source, request, row, f)).ToArray()));
                groupKeys.Add(new ValueKey(dimensions.Select(f => ReadValue(source, request, row, f)).ToArray()));
                CheckOutputSize(request, rowKeys.Count, columnKeys.Count);
                if ((long)groupKeys.Count * request.Measures.Length > MaxGroups)
                    throw new InvalidOperationException("The pivot has too many groups and values. Add filters, remove values, or choose fields with fewer distinct values.");
                matchingIndexes.Add(i);
            }

            var fieldIndexes = source.Fields.ToDictionary(f => f.Key, f => f.Index);
            var pivots = new PivotTable[request.Measures.Length];
            long totalGroups = 0;
            for (int measureIndex = 0; measureIndex < request.Measures.Length; measureIndex++)
            {
                cancellation.ThrowIfCancellationRequested();
                var measure = request.Measures[measureIndex];
                // A separate value key prevents numeric conversion from changing a grouping key
                // when the same field is used both as a dimension and as a measure.
                const string valueKey = "measureValue";
                var factory = CreateFactory(measure.Aggregation, valueKey);
                var cube = new NReco.PivotData.PivotData(dimensions.Select(i => source.Fields[i].Key).ToArray(), factory, true);
                cube.LazyAdd = false;

                IEnumerable<string[]> ReadRows()
                {
                    foreach (int index in matchingIndexes)
                    {
                        cancellation.ThrowIfCancellationRequested();
                        yield return source.Rows[index];
                        if (totalGroups + cube.Count > MaxGroups)
                            throw new InvalidOperationException("The pivot has too many groups and values. Add filters, remove values, or choose fields with fewer distinct values.");
                    }
                }

                cube.ProcessData(ReadRows(), (record, key) => key == valueKey
                    ? ReadMeasure(source, request, (string[])record, measure)
                    : ReadValue(source, request, (string[])record, fieldIndexes[key]));
                totalGroups += cube.Count;
                cancellation.ThrowIfCancellationRequested();
                pivots[measureIndex] = new PivotTable(request.Rows.Select(i => source.Fields[i].Key).ToArray(),
                    request.Columns.Select(i => source.Fields[i].Key).ToArray(), cube);
            }

            var pivot = pivots[0];
            var displayedRowKeys = request.Rows.Length == 0 ? new ValueKey[0] : pivot.RowKeys;
            var displayedColumnKeys = request.Columns.Length == 0 ? new ValueKey[0] : pivot.ColumnKeys;
            CheckOutputSize(request, displayedRowKeys.Length, displayedColumnKeys.Length);
            var table = new DataTable { Locale = CultureInfo.InvariantCulture };
            void AddColumn(string caption) => table.Columns.Add(new DataColumn("c" + table.Columns.Count, typeof(object)) { Caption = caption });
            if (request.Rows.Length == 0) AddColumn("Rows");
            else foreach (var i in request.Rows) AddColumn(source.Fields[i].Label);
            foreach (var key in displayedColumnKeys)
            {
                string groupLabel = string.Join(" | ", key.DimKeys.Select((value, i) =>
                    source.Fields[request.Columns[i]].Label + " = " + Label(value)));
                foreach (var measure in request.Measures)
                    AddColumn(groupLabel + " | " + measure.GetLabel(source));
            }
            foreach (var measure in request.Measures)
                AddColumn((request.Columns.Length == 0 ? "" : "Grand total | ") + measure.GetLabel(source));

            void AddRow(int? rowIndex)
            {
                cancellation.ThrowIfCancellationRequested();
                var cells = new object[table.Columns.Count];
                int offset = Math.Max(1, request.Rows.Length);
                if (rowIndex.HasValue)
                    Array.Copy(pivot.RowKeys[rowIndex.Value].DimKeys, cells, request.Rows.Length);
                else
                {
                    cells[0] = "Grand total";
                    for (int i = 1; i < offset; i++) cells[i] = "";
                }
                int outputColumn = offset;
                for (int column = 0; column < displayedColumnKeys.Length; column++)
                {
                    for (int measure = 0; measure < pivots.Length; measure++)
                    {
                        cancellation.ThrowIfCancellationRequested();
                        cells[outputColumn++] = pivots[measure][rowIndex, column].Value;
                    }
                }
                for (int measure = 0; measure < pivots.Length; measure++)
                {
                    cancellation.ThrowIfCancellationRequested();
                    cells[outputColumn++] = pivots[measure][rowIndex, null].Value;
                }
                table.Rows.Add(cells);
            }
            for (int row = 0; row < displayedRowKeys.Length; row++) AddRow(row);
            AddRow(null);
            return new PivotResult(table, matchingIndexes.Count, source, request, displayedRowKeys, displayedColumnKeys);
        }

        private static void CheckOutputSize(PivotRequest request, int rowGroups, int columnGroups)
        {
            long columns = Math.Max(1, request.Rows.Length) + (long)request.Measures.Length *
                (request.Columns.Length == 0 ? 1 : columnGroups + 1);
            long rows = request.Rows.Length == 0 ? 1 : (long)rowGroups + 1;
            if (columns > MaxColumns || rows * columns > MaxCells)
                throw new InvalidOperationException("The pivot is too large. Add filters, remove values, or choose fields with fewer distinct values.");
        }

        internal static object ReadValue(PivotSnapshot source, PivotRequest request, string[] row, int index)
        {
            var text = row[index];
            return NormalizeNull(source.Fields[index], request, text) ?? (object)DBNull.Value;
        }

        private static string NormalizeNull(PivotField field, PivotRequest request, string text) =>
            text == null || ((request.NullTextIsNull || field.IsNumeric) && text == "NULL") ? null : text;

        private static object ReadMeasure(PivotSnapshot source, PivotRequest request, string[] row, PivotMeasure measure)
        {
            var value = ReadValue(source, request, row, measure.Value);
            if (value == DBNull.Value || !RequiresNumber(measure.Aggregation)) return value;
            decimal number;
            // Fail explicitly instead of letting NReco silently ignore unsupported numeric values.
            if (!TryParseDecimalExact((string)value, out number) || number == decimal.MinValue)
                throw new InvalidOperationException("A value in " + source.Fields[measure.Value].Label +
                    " cannot be aggregated as a decimal. Cast or round this column in SQL first.");
            return number;
        }

        internal static Func<string[], bool> CreateFilterMatcher(PivotSnapshot source, PivotRequest request)
        {
            var predicates = new List<Func<string[], bool>>();
            foreach (var filter in request.Filters)
            {
                if (filter == null || filter.Field < 0 || filter.Field >= source.Fields.Length)
                    throw new InvalidOperationException("Choose a valid field for every filter.");
                if (!Enum.IsDefined(typeof(PivotFilterOperator), filter.Operator))
                    throw new InvalidOperationException("Choose a valid filter operator.");
                int index = filter.Field;
                var field = source.Fields[index];
                Func<string[], string> read = row => NormalizeNull(field, request, row[index]);
                if (filter.Operator == PivotFilterOperator.IsNull)
                {
                    predicates.Add(row => read(row) == null);
                    continue;
                }
                if (filter.Operator == PivotFilterOperator.IsNotNull)
                {
                    predicates.Add(row => read(row) != null);
                    continue;
                }
                if (filter.Operator == PivotFilterOperator.In)
                {
                    if (filter.Values == null || filter.Values.Length == 0)
                        throw new InvalidOperationException("Choose values for the " + field.Label + " filter.");
                    // Selected values identify the exact displayed groups, including case and null.
                    var selected = new HashSet<string>(filter.Values.Select(v => NormalizeNull(field, request, v)), StringComparer.Ordinal);
                    predicates.Add(row => selected.Contains(read(row)));
                    continue;
                }
                if (filter.Operator == PivotFilterOperator.Contains)
                {
                    string term = filter.Text ?? "";
                    predicates.Add(row => read(row)?.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0);
                    continue;
                }

                bool range = filter.Operator >= PivotFilterOperator.GreaterThan;
                if (range && !field.IsNumeric)
                    throw new InvalidOperationException("The comparison filter for " + field.Label + " requires a numeric field.");
                string criterion = NormalizeNull(field, request, filter.Text);
                if (criterion == null)
                {
                    if (range) throw new InvalidOperationException("Enter a number for the " + field.Label + " filter.");
                    bool equalsNull = filter.Operator == PivotFilterOperator.Equals;
                    predicates.Add(row => (read(row) == null) == equalsNull);
                    continue;
                }

                var comparisonOperator = filter.Operator;
                if (field.IsNumeric)
                {
                    decimal criterionNumber;
                    if (!TryParseDecimalExact(criterion, out criterionNumber))
                        throw new InvalidOperationException("Enter a number within decimal precision for the " + field.Label + " filter, using a period for decimals.");
                    predicates.Add(row =>
                    {
                        string value = read(row);
                        if (value == null) return false;
                        decimal number;
                        if (!TryParseDecimalExact(value, out number))
                            throw new InvalidOperationException("A value in " + field.Label + " cannot be compared as a decimal. Cast or round this column in SQL first.");
                        return Compare(number.CompareTo(criterionNumber), comparisonOperator);
                    });
                }
                else
                {
                    predicates.Add(row =>
                    {
                        string value = read(row);
                        return value != null && Compare(StringComparer.OrdinalIgnoreCase.Compare(value, criterion), comparisonOperator);
                    });
                }
            }
            return row =>
            {
                foreach (var predicate in predicates)
                    if (!predicate(row)) return false;
                return true;
            };
        }

        private static bool Compare(int comparison, PivotFilterOperator operation)
        {
            switch (operation)
            {
                case PivotFilterOperator.Equals: return comparison == 0;
                case PivotFilterOperator.NotEquals: return comparison != 0;
                case PivotFilterOperator.GreaterThan: return comparison > 0;
                case PivotFilterOperator.GreaterThanOrEqual: return comparison >= 0;
                case PivotFilterOperator.LessThan: return comparison < 0;
                case PivotFilterOperator.LessThanOrEqual: return comparison <= 0;
                default: throw new InvalidOperationException("Choose a valid comparison operator.");
            }
        }

        internal static bool TryParseDecimalExact(string text, out decimal number)
        {
            if (!decimal.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) return false;
            // TryParse can succeed after rounding SQL decimal(38,s) or underflowing to zero.
            // Compare canonical decimal text so neither filtering nor aggregation silently
            // changes the source value. Equivalent scale and exponent notation are accepted.
            string input = CanonicalNumber(text);
            return input != null && input == CanonicalNumber(number.ToString(CultureInfo.InvariantCulture));
        }

        private static string CanonicalNumber(string text)
        {
            text = text.Trim();
            bool negative = text[0] == '-';
            if (negative || text[0] == '+') text = text.Substring(1);
            int exponentOffset = text.IndexOfAny(new[] { 'e', 'E' });
            string mantissa = exponentOffset < 0 ? text : text.Substring(0, exponentOffset);
            int point = mantissa.IndexOf('.');
            int fractionDigits = point < 0 ? 0 : mantissa.Length - point - 1;
            string digits = mantissa.Replace(".", "").TrimStart('0');
            if (digits.Length == 0) return "0";
            long exponent = 0;
            if (exponentOffset >= 0 && !long.TryParse(text.Substring(exponentOffset + 1),
                NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out exponent)) return null;
            string significant = digits.TrimEnd('0');
            try { exponent = checked(exponent - fractionDigits + digits.Length - significant.Length); }
            catch (OverflowException) { return null; }
            return (negative ? "-" : "") + significant + ":" + exponent.ToString(CultureInfo.InvariantCulture);
        }

        internal static string DescribeFilter(PivotSnapshot source, PivotFilter filter)
        {
            string field = source.Fields[filter.Field].Label;
            if (filter.Operator == PivotFilterOperator.IsNull) return field + " is null";
            if (filter.Operator == PivotFilterOperator.IsNotNull) return field + " is not null";
            if (filter.Operator == PivotFilterOperator.In)
                return field + " is one of " + string.Join(", ", filter.Values.Take(10).Select(v => Label(v))) +
                    (filter.Values.Length > 10 ? " (and " + (filter.Values.Length - 10) + " more)" : "");
            string operation;
            switch (filter.Operator)
            {
                case PivotFilterOperator.Contains: operation = "contains"; break;
                case PivotFilterOperator.Equals: operation = "="; break;
                case PivotFilterOperator.NotEquals: operation = "<>"; break;
                case PivotFilterOperator.GreaterThan: operation = ">"; break;
                case PivotFilterOperator.GreaterThanOrEqual: operation = ">="; break;
                case PivotFilterOperator.LessThan: operation = "<"; break;
                case PivotFilterOperator.LessThanOrEqual: operation = "<="; break;
                default: operation = filter.Operator.ToString(); break;
            }
            return field + " " + operation + " " + Label(filter.Text);
        }

        internal static string Label(object value) => value == null || value == DBNull.Value ? "(NULL)" :
            "'" + Convert.ToString(value, CultureInfo.InvariantCulture).Replace("'", "''") + "'";

        private static IAggregatorFactory CreateFactory(PivotAggregation aggregation, string field)
        {
            switch (aggregation)
            {
                case PivotAggregation.CountRows: return new CountAggregatorFactory();
                case PivotAggregation.CountValues: return new CountAggregatorFactory(field);
                case PivotAggregation.DistinctCount: return new CountUniqueAggregatorFactory(field);
                case PivotAggregation.Sum: return new SumAggregatorFactory(field);
                case PivotAggregation.Average: return new AverageAggregatorFactory(field);
                case PivotAggregation.Minimum: return new MinAggregatorFactory(field);
                case PivotAggregation.Maximum: return new MaxAggregatorFactory(field);
                default: throw new ArgumentOutOfRangeException(nameof(aggregation));
            }
        }
    }
}
