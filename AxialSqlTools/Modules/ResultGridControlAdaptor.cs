using Microsoft.SqlServer.Management.UI.Grid;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Web.UI.Design;

namespace AxialSqlTools
{

    public static class DataTableExtensions
    {
        public static string ToCsv(this DataTable dt)
        {
            StringBuilder sb = new StringBuilder();

            IEnumerable<string> columnNames = dt.Columns
                .Cast<DataColumn>()
                .Select(column => column.ColumnName);

            sb.AppendLine(string.Join(",", columnNames));
            foreach (DataRow row in dt.Rows)
            {
                IEnumerable<string> fields = row.ItemArray.Select(field =>
                  string.Concat("\"", field.ToString().Replace("\"", "\"\""), "\""));
                sb.AppendLine(string.Join(",", fields));
            }

            return sb.ToString();
        }
    }

    public static class EnumerableExtension
    {
        public static IEnumerable<TResult> ZipIt<TSource, TResult>(this IEnumerable<IEnumerable<TSource>> collection,
                                            Func<IEnumerable<TSource>, TResult> resultSelector)
        {
            var enumerators = collection.Select(c => c.GetEnumerator()).ToList();
            while (enumerators.All(e => e.MoveNext()))
            {
                yield return resultSelector(enumerators.Select(e => e.Current).ToList());
            }
        }
    }

    public class ResultGridSelectedCell
    {
        public int ColumnIndex { get; }
        public long RowIndex { get; }

        public ResultGridSelectedCell(long nRowIndex, int nColIndex)
        {
            RowIndex = nRowIndex;
            ColumnIndex = nColIndex;
        }

        public static implicit operator ResultGridSelectedCell(BlockOfCells blockOfCells)
        {
            return new ResultGridSelectedCell(blockOfCells.OriginalY, blockOfCells.OriginalX);
        }
    }

    public class ResultGridControlAdaptor : IDisposable
    {
        public bool IsDisposed { get; private set; }
        public long RowCount { get; private set; }
        public long ColumnCount { get; private set; }

        private readonly IGridControl _gridControl;

        public ResultGridControlAdaptor(IGridControl gridControl)
        {
            _gridControl = gridControl ?? throw new ArgumentNullException(nameof(gridControl));
            ColumnCount = gridControl.ColumnsNumber;
            RowCount = gridControl?.GridStorage?.NumRows() ?? 0;
        }

        public DataTable GridFocusAsDatatable()
        {
            // DEFINE COLUMN
            var datatable = SchemaResultGrid();

            // DEFINE ROWS
            for (var nRowIndex = 0L; nRowIndex < RowCount; ++nRowIndex)
            {
                var rows = new List<object>();
                for (var nColIndex = 1; nColIndex < ColumnCount; nColIndex++)
                {
                    var cellText = _gridControl.GridStorage.GetCellDataAsString(nRowIndex, nColIndex);

                    if (cellText == "NULL")
                        rows.Add(null);
                    else
                    {
                        var column = datatable.Columns[nColIndex - 1];

                        if (column.DataType == typeof(bool))
                            cellText = cellText == "0" ? "False" : "True";

                        if (column.DataType == typeof(Guid))
                            rows.Add(new Guid(cellText));
                        else if (column.DataType == typeof(DateTime) || column.DataType == typeof(DateTimeOffset))
                            rows.Add(DateTime.Parse(cellText));
                        else if (column.DataType == typeof(byte[]))
                            rows.Add(Encoding.UTF8.GetBytes(cellText));
                        else
                        {
                            var typedValue = Convert.ChangeType(cellText, column.DataType, CultureInfo.InvariantCulture);
                            rows.Add(typedValue);
                        }
                    }
                }

                datatable.Rows.Add(rows.ToArray());
            }

            datatable.AcceptChanges();
            return datatable;
        }

        public DataTable SchemaResultGrid()
        {
            var datatable = new DataTable();
            var schemaTable = GridAccess.GetNonPublicField(_gridControl.GridStorage, "m_schemaTable") as DataTable;
            for (var column = 1; column < ColumnCount; column++)
            {
                var columnType = (Type)schemaTable.Rows[column - 1][12];
                var columnText = GetColumnName(column);
                datatable.Columns.Add(columnText, columnType);
            }
            return datatable;
        }

        public object GetCellValueAsString(long nRowIndex, int nColIndex)
        {
            var cellText = _gridControl.GridStorage.GetCellDataAsString(nRowIndex, nColIndex) ?? "";

            if (cellText.Equals("NULL", StringComparison.OrdinalIgnoreCase))
                return cellText;

            var schema = GridAccess.GetNonPublicField(_gridControl.GridStorage, "m_schemaTable") as DataTable;
            var columnType = (Type)schema.Rows[nColIndex - 1][12];

            if (columnType == typeof(bool))
                return cellText == "1" ? 1 : 0;

            if (columnType == typeof(Guid))
                return string.Format("'{0}'", cellText);

            if (columnType == typeof(int) || columnType == typeof(decimal) || columnType == typeof(long) || columnType == typeof(double) || columnType == typeof(float) || columnType == typeof(byte))
                return Convert.ChangeType(cellText, columnType, CultureInfo.InvariantCulture);

            if (columnType == typeof(Guid) || columnType == typeof(DateTime) || columnType == typeof(DateTimeOffset) || columnType == typeof(byte[]))
            {
                columnType = typeof(string);
                var @values = Convert.ChangeType(cellText, columnType, CultureInfo.InvariantCulture);
                return string.Format("N'{0}'", @values);
            }

            return string.Format("N'{0}'", cellText.Replace("'", "''"));
        }

        public IEnumerable<(Type, string)> GetColumnTypes()
        {
            var result = new List<(Type, string)>();

            return result;
        }

        public string GetColumnName(int nColIndex)
        {
            if (nColIndex > 0 && nColIndex < ColumnCount)
            {
                _gridControl.GetHeaderInfo(nColIndex, out var columnName, out Bitmap _);
                return columnName;
            }

            return string.Empty;
        }

        public (Type, string) GetColumnType(int nColIndex)
        {
            if (nColIndex > 0 && nColIndex < ColumnCount)
            {
                var schema = GridAccess.GetNonPublicField(_gridControl.GridStorage, "m_schemaTable") as DataTable;
                var columnType = (Type)schema.Rows[nColIndex - 1][12];
                var columnText = GetColumnName(nColIndex);
                return (columnType, columnText);
            }

            return (default, string.Empty);
        }

        public string[] GetBracketColumns()
        {
            var columnHeaders = new string[ColumnCount - 1];
            for (var colIndex = 1; colIndex < ColumnCount; colIndex++)
            {
                var columnName = GetColumnName(colIndex);
                if (columnHeaders.Contains("[" + columnName + "]"))
                    columnName = columnName + "_" + colIndex.ToString();

                columnHeaders[colIndex - 1] = !columnName.StartsWith("[") ? "[" + columnName + "]" : columnName;
            }

            return columnHeaders;
        }

        public IEnumerable<IEnumerable<object>> GridAsQuerySql()
        {
            var rows = new List<List<object>>();
            for (var nRowIndex = 0L; nRowIndex < RowCount; ++nRowIndex)
            {
                var columns = new List<object>();
                for (var nColIndex = 1; nColIndex < ColumnCount; nColIndex++)
                {
                    var cellText = GetCellValueAsString(nRowIndex, nColIndex);
                    columns.Add(cellText);
                }

                rows.Add(columns);
            }

            return rows;
        }


        public IEnumerable<string> GridSelectedAsQuerySql()
        {
            var resultGridSelected = GridSelectedAsDictionary();
            var columnJoins = string.Join(", ", resultGridSelected.Select(q => q.Key.StartsWith("[") ? q.Key : "[" + q.Key + "]"));
            var rowJoins = resultGridSelected.Select(q => q.Value).ZipIt(xs => string.Join(", ", xs));
            var linkedList = new LinkedList<string>(rowJoins);
            linkedList.AddFirst(columnJoins);
            return linkedList;
        }

        public DataTable GridSelectedAsDataTable()
        {
            var datatable = new DataTable();

            // COLUMNS
            var schemaTable = GridAccess.GetNonPublicField(_gridControl.GridStorage, "m_schemaTable") as DataTable;
            foreach (BlockOfCells cell in _gridControl.SelectedCells)
            {
                for (var col = cell.X; col <= cell.Right; col++)
                {
                    var columnType = (Type)schemaTable.Rows[col - 1][12];
                    var columnText = GetColumnName(col);
                    datatable.Columns.Add(columnText, columnType);
                }
            }

            // ROWS
            foreach (BlockOfCells cell in _gridControl.SelectedCells)
            {
                for (var row = cell.Y; row <= cell.Bottom; row++)
                {
                    var rows = new List<object>();
                    var nColIndex = 0;
                    for (var col = cell.X; col <= cell.Right; col++)
                    {
                        var cellText = _gridControl.GridStorage.GetCellDataAsString(row, col);

                        if (cellText == "NULL")
                            rows.Add(null);
                        else
                        {
                            var column = datatable.Columns[nColIndex];
                            if (column.DataType == typeof(bool))
                                cellText = cellText == "0" ? "False" : "True";

                            if (column.DataType == typeof(Guid))
                                rows.Add(new Guid(cellText));
                            else if (column.DataType == typeof(DateTime) || column.DataType == typeof(DateTimeOffset))
                                rows.Add(DateTime.Parse(cellText));
                            else if (column.DataType == typeof(byte[]))
                                rows.Add(Encoding.UTF8.GetBytes(cellText));
                            else
                            {
                                var typedValue = Convert.ChangeType(cellText, column.DataType, CultureInfo.InvariantCulture);
                                rows.Add(typedValue);
                            }
                        }

                        nColIndex++;
                    }
                    datatable.Rows.Add(rows.ToArray());
                }
            }

            datatable.AcceptChanges();
            return datatable;
        }

        public Dictionary<string, List<object>> GridSelectedAsDictionary()
        {
            var gridHeader = GridAccess.GetNonPublicField(_gridControl, "m_gridHeader") as GridHeader;
            var gridResult = new Dictionary<string, List<object>>();
            foreach (BlockOfCells cell in _gridControl.SelectedCells)
            {
                for (var row = cell.Y; row <= cell.Bottom; row++)
                {
                    for (var col = cell.X; col <= cell.Right; col++)
                    {
                        string column = gridHeader[col].Text.Trim();
                        if (column.StartsWith("<"))
                        {
                            try
                            {
                                var cols = column.Split('(');
                                if (cols.Length > 1)
                                    column = cols[1].TrimEnd(")>".ToArray());
                            }
                            catch { }
                        }

                        if (!gridResult.ContainsKey(column))
                            gridResult.Add(column, new List<object>());

                        var cellText = GetCellValueAsString(row, col);
                        gridResult[column].Add(cellText);
                    }
                }
            }

            return gridResult;
        }
        /*
        public void SetColumnBackground(int columnIndex, Color backgroundColor)
        {
            var columnsCollection = _gridControl.GetBaseTypeField<GridColumnCollection>("m_columns");
            columnsCollection[columnIndex].BackgroundBrush.Color = backgroundColor;
        }

        public void SetRangeColumnBackground(int beginColumn, int endColumn, Color backgroundColor)
        {
            var columnsCollection = _gridControl.GetBaseTypeField<GridColumnCollection>("m_columns");
            for (var i = beginColumn; i < endColumn; i++)
            {
                columnsCollection[i].BackgroundBrush.Color = backgroundColor;
            }
        }*/

        public IEnumerable<ResultGridSelectedCell> GetSelectedCells()
        {
            return _gridControl.SelectedCells
                .Cast<BlockOfCells>()
                .Select((Func<BlockOfCells, ResultGridSelectedCell>)((item) => item));
        }

        public void Dispose()
        {
            if (IsDisposed)
                return;

            GC.ReRegisterForFinalize(ColumnCount);
            GC.ReRegisterForFinalize(RowCount);
            GC.ReRegisterForFinalize(_gridControl);
            GC.ReRegisterForFinalize(this);

            IsDisposed = true;
        }
    }
}