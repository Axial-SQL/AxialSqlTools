using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.VisualStudio.Shell;
using System;
using System.Data;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace AxialSqlTools.PivotGrid
{
    internal static class PivotGridSnapshot
    {
        // Bound additional memory inside the SSMS process. Never silently sample a result.
        internal const int MaxRows = 1000000;
        internal const int MaxCells = 5000000;
        internal const long MaxCharacters = 64000000;

        public static async Task<PivotSnapshot> CaptureAsync(IGridControl grid, CancellationToken cancellation,
            Action<int, int> progress)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellation);
            var storage = grid.GridStorage;
            long rowCount = storage.NumRows();
            int columnCount = grid.ColumnsNumber - 1; // The first grid column contains row numbers.
            if (rowCount <= 0 || columnCount <= 0)
                throw new InvalidOperationException("The result grid is empty. Run a query first.");
            if (rowCount > MaxRows || rowCount * columnCount > MaxCells)
                throw new InvalidOperationException("This grid is too large to snapshot. Reduce the query to at most 1,000,000 rows and 5,000,000 cells.");
            var schema = GridAccess.GetNonPublicField(storage, "m_schemaTable") as DataTable;
            if (schema == null || schema.Rows.Count != columnCount || !schema.Columns.Contains("DataType"))
                throw new InvalidOperationException("SSMS did not expose the result column types. Finish the query and try again.");

            var fields = new PivotField[columnCount];
            for (int column = 0; column < columnCount; column++)
            {
                grid.GetHeaderInfo(column + 1, out var name, out System.Drawing.Bitmap _);
                fields[column] = new PivotField(column, name, schema.Rows[column]["DataType"] as Type);
            }
            var rows = new string[(int)rowCount][];
            long characters = 0;
            var slice = Stopwatch.StartNew();
            for (int row = 0; row < rows.Length; row++)
            {
                cancellation.ThrowIfCancellationRequested();
                rows[row] = new string[columnCount];
                for (int column = 0; column < columnCount; column++)
                {
                    string value = storage.GetCellDataAsString(row, column + 1);
                    rows[row][column] = value;
                    characters += value?.Length ?? 0;
                    if (characters > MaxCharacters)
                        throw new InvalidOperationException("The grid contains too much text to snapshot. Select fewer rows or columns in SQL.");
                }
                if (slice.ElapsedMilliseconds >= 25)
                {
                    progress(row + 1, rows.Length);
                    // Grid APIs must stay on the UI thread; yield regularly for paint and Cancel.
                    await Dispatcher.Yield(DispatcherPriority.Background);
                    cancellation.ThrowIfCancellationRequested();
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellation);
                    if (!ReferenceEquals(grid.GridStorage, storage) || grid.ColumnsNumber != columnCount + 1 ||
                        storage.NumRows() != rowCount || !ReferenceEquals(GridAccess.GetNonPublicField(storage, "m_schemaTable"), schema))
                        throw new InvalidOperationException("The result grid changed while copying. Wait for the query to finish and open Pivot Grid again.");
                    slice.Restart();
                }
            }
            return new PivotSnapshot(fields, rows);
        }
    }
}
