using System;
using System.Data;

namespace AxialSqlTools
{
    internal static class QuickSearchNavigationTarget
    {
        public static ScriptObjectSelectionItem FromResult(DataRow row)
        {
            if (row == null) throw new ArgumentNullException(nameof(row));
            // Use catalog identifiers, never the preview or the decorated job-step caption.
            // The exact type also distinguishes scalar and table-valued function folders.
            return new ScriptObjectSelectionItem(
                Required(row, "NavigationTypeDesc"),
                Required(row, "ScriptSchemaName"),
                Required(row, "ScriptObjectName"), 0,
                Required(row, "ScriptDatabaseName"), 0, null);
        }

        private static string Required(DataRow row, string column)
        {
            string value = row.Table.Columns.Contains(column) ? row[column] as string : null;
            if (string.IsNullOrEmpty(value))
                throw new InvalidOperationException("This result has no navigation target. Run the search again.");
            return value;
        }
    }
}
