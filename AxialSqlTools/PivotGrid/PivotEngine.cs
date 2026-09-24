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
        public string Label { get; }
        public bool IsNumeric { get; }

        public PivotField(int index, string name, Type type)
        {
            Index = index;
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

    internal sealed class PivotRequest
    {
        public int[] Rows { get; set; } = new int[0];
        public int[] Columns { get; set; } = new int[0];
        public int Value { get; set; } = -1;
        public PivotAggregation Aggregation { get; set; }
        public bool NullTextIsNull { get; set; } = true;
        public int FilterField { get; set; } = -1;
        public string FilterText { get; set; } = "";

        internal PivotRequest Copy() => new PivotRequest
        {
            Rows = (int[])Rows.Clone(), Columns = (int[])Columns.Clone(), Value = Value,
            Aggregation = Aggregation, NullTextIsNull = NullTextIsNull,
            FilterField = FilterField, FilterText = FilterText
        };
    }

    internal sealed class PivotResult
    {
        public DataTable Table { get; }
        public int MatchedRows { get; }
        private readonly PivotSnapshot source;
        private readonly PivotRequest request;
        private readonly object[][] rowKeys, columnKeys;

        public PivotResult(DataTable table, int matchedRows, PivotSnapshot source, PivotRequest request,
            ValueKey[] rowKeys, ValueKey[] columnKeys)
        {
            Table = table; MatchedRows = matchedRows; this.source = source; this.request = request;
            this.rowKeys = rowKeys.Select(k => (object[])k.DimKeys.Clone()).ToArray();
            this.columnKeys = columnKeys.Select(k => (object[])k.DimKeys.Clone()).ToArray();
        }

        public bool CanDrillDown(int row, int column) => row >= 0 && row < Table.Rows.Count &&
            column >= Math.Max(1, request.Rows.Length) && column < Table.Columns.Count;

        public PivotDetails GetUnderlyingRows(int row, int column, CancellationToken cancellation)
        {
            cancellation.ThrowIfCancellationRequested();
            if (!CanDrillDown(row, column)) throw new ArgumentOutOfRangeException(nameof(column), "Select a value or total cell.");
            var rowKey = row < rowKeys.Length ? rowKeys[row] : null;
            int pivotColumn = column - Math.Max(1, request.Rows.Length);
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
                if (PivotEngine.MatchesFilter(request, record) && Matches(record, request.Rows, rowKey) &&
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
            if (request.FilterField >= 0)
                labels.Add("Filter: " + source.Fields[request.FilterField].Label + " contains '" + request.FilterText + "'");
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

        public static PivotResult Build(PivotSnapshot source, PivotRequest request, CancellationToken cancellation)
        {
            cancellation.ThrowIfCancellationRequested();
            // Keep drill-down tied to the configuration that produced the displayed result.
            request = request.Copy();
            var dimensions = request.Rows.Concat(request.Columns).ToArray();
            if (dimensions.Distinct().Count() != dimensions.Length)
                throw new InvalidOperationException("Choose different fields for Rows and Columns.");
            if (dimensions.Length > 6)
                throw new InvalidOperationException("Choose up to six grouping fields in total.");
            if (dimensions.Any(i => i < 0 || i >= source.Fields.Length) ||
                request.FilterField < -1 || request.FilterField >= source.Fields.Length)
                throw new InvalidOperationException("Select valid grid fields.");
            if (request.Aggregation != PivotAggregation.CountRows &&
                (request.Value < 0 || request.Value >= source.Fields.Length))
                throw new InvalidOperationException("Choose a Values field.");
            if (RequiresNumber(request.Aggregation) && !source.Fields[request.Value].IsNumeric)
                throw new InvalidOperationException("This aggregation requires a numeric Values field.");

            var fieldIndexes = source.Fields.ToDictionary(f => f.Key, f => f.Index);
            string valueKey = request.Value >= 0 ? source.Fields[request.Value].Key : null;
            var factory = CreateFactory(request.Aggregation, valueKey);
            var cube = new NReco.PivotData.PivotData(dimensions.Select(i => source.Fields[i].Key).ToArray(), factory, true);
            cube.LazyAdd = false;
            var rowKeys = new HashSet<ValueKey>();
            var columnKeys = new HashSet<ValueKey>();
            int matchedRows = 0;

            object ReadCell(string[] row, int index)
            {
                var text = row[index];
                return text == null || ((request.NullTextIsNull || source.Fields[index].IsNumeric) && text == "NULL") ? DBNull.Value : (object)text;
            }

            IEnumerable<string[]> ReadRows()
            {
                foreach (var row in source.Rows)
                {
                    cancellation.ThrowIfCancellationRequested();
                    if (!MatchesFilter(request, row)) continue;
                    rowKeys.Add(new ValueKey(request.Rows.Select(i => ReadCell(row, i)).ToArray()));
                    columnKeys.Add(new ValueKey(request.Columns.Select(i => ReadCell(row, i)).ToArray()));
                    if (columnKeys.Count + 1 + Math.Max(1, request.Rows.Length) > MaxColumns ||
                        (long)(rowKeys.Count + 1) * (columnKeys.Count + 1 + Math.Max(1, request.Rows.Length)) > MaxCells)
                        throw new InvalidOperationException("The pivot is too large. Filter the data or choose fields with fewer distinct values.");
                    matchedRows++;
                    yield return row;
                    if (cube.Count > MaxGroups)
                        throw new InvalidOperationException("The pivot has too many groups. Filter the data or choose fields with fewer distinct values.");
                }
            }

            cube.ProcessData(ReadRows(), (record, key) => ReadValue(source, request, (string[])record, fieldIndexes[key]));
            cancellation.ThrowIfCancellationRequested();

            var pivot = new PivotTable(request.Rows.Select(i => source.Fields[i].Key).ToArray(),
                request.Columns.Select(i => source.Fields[i].Key).ToArray(), cube);
            var table = new DataTable { Locale = CultureInfo.InvariantCulture };
            void AddColumn(string caption) => table.Columns.Add(new DataColumn("c" + table.Columns.Count, typeof(object)) { Caption = caption });
            if (request.Rows.Length == 0) AddColumn("Rows");
            else foreach (var i in request.Rows) AddColumn(source.Fields[i].Label);
            foreach (var key in pivot.ColumnKeys)
                AddColumn(string.Join(" | ", key.DimKeys.Select((value, i) =>
                    source.Fields[request.Columns[i]].Label + " = " + Label(value))));
            AddColumn(request.Columns.Length == 0 ? "Value" : "Grand total");

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
                for (int column = 0; column < pivot.ColumnKeys.Length; column++)
                {
                    cancellation.ThrowIfCancellationRequested();
                    cells[offset + column] = pivot[rowIndex, column].Value;
                }
                cells[cells.Length - 1] = pivot[rowIndex, null].Value;
                table.Rows.Add(cells);
            }
            for (int row = 0; row < pivot.RowKeys.Length; row++) AddRow(row);
            AddRow(null);
            return new PivotResult(table, matchedRows, source, request, pivot.RowKeys, pivot.ColumnKeys);
        }

        internal static bool MatchesFilter(PivotRequest request, string[] row) => request.FilterField < 0 ||
            (row[request.FilterField] ?? "").IndexOf(request.FilterText ?? "", StringComparison.OrdinalIgnoreCase) >= 0;

        internal static object ReadValue(PivotSnapshot source, PivotRequest request, string[] row, int index)
        {
            var text = row[index];
            if (text == null || ((request.NullTextIsNull || source.Fields[index].IsNumeric) && text == "NULL")) return DBNull.Value;
            if (index != request.Value || !RequiresNumber(request.Aggregation)) return text;
            decimal number;
            // Fail explicitly instead of letting NReco silently ignore unsupported numeric values.
            if (!decimal.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number) || number == decimal.MinValue)
                throw new InvalidOperationException("A value in " + source.Fields[index].Label +
                    " cannot be aggregated as a decimal. Cast or round this column in SQL first.");
            return number;
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
