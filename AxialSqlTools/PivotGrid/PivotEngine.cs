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
    }

    internal sealed class PivotResult
    {
        public DataTable Table { get; }
        public int MatchedRows { get; }
        public PivotResult(DataTable table, int matchedRows) { Table = table; MatchedRows = matchedRows; }
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
                    if (request.FilterField >= 0 && (row[request.FilterField] ?? "").IndexOf(
                        request.FilterText ?? "", StringComparison.OrdinalIgnoreCase) < 0) continue;
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

            cube.ProcessData(ReadRows(), (record, key) =>
            {
                var cell = ReadCell((string[])record, fieldIndexes[key]);
                if (key != valueKey || !RequiresNumber(request.Aggregation) || cell == DBNull.Value) return cell;
                decimal number;
                // Fail explicitly instead of letting NReco silently ignore values outside Decimal's range.
                if (!decimal.TryParse((string)cell, NumberStyles.Float, CultureInfo.InvariantCulture, out number) ||
                    number == decimal.MinValue)
                    throw new InvalidOperationException("A value in " + source.Fields[request.Value].Label +
                        " cannot be aggregated as a decimal. Cast or round this column in SQL first.");
                return number;
            });
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
            return new PivotResult(table, matchedRows);
        }

        private static string Label(object value) => value == null || value == DBNull.Value ? "(NULL)" :
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
