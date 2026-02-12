using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace AxialSqlTools
{
    public partial class QuickSearchWindowControl : UserControl
    {
        private string selectedConnectionString;
        private string selectedDatabase;
        private string selectedServer;

        public QuickSearchWindowControl()
        {
            this.InitializeComponent();
        }

        private void Button_SelectConnection_Click(object sender, RoutedEventArgs e)
        {
            var ci = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer();
            if (ci == null)
            {
                MessageBox.Show("Please select a server or database node in Object Explorer first.", "Quick Search");
                return;
            }

            selectedConnectionString = ci.FullConnectionString;
            selectedDatabase = ci.Database;
            selectedServer = ci.ServerName;
            Label_ConnectionDescription.Content = $"Server: [{selectedServer}] / Database: [{selectedDatabase}]";
        }

        private async void Button_Search_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(selectedConnectionString))
            {
                MessageBox.Show("Select a connection from Object Explorer first.", "Quick Search");
                return;
            }

            string searchText = TextBox_SearchText.Text?.Trim();
            if (string.IsNullOrEmpty(searchText))
            {
                MessageBox.Show("Enter text to search.", "Quick Search");
                return;
            }

            if (!AnyTypeSelected())
            {
                MessageBox.Show("Select at least one object type.", "Quick Search");
                return;
            }

            try
            {
                Button_Search.IsEnabled = false;
                TextBlock_ResultCount.Text = "Searching...";

                bool allDatabases = CheckBox_AllDatabases.IsChecked == true;
                bool wholeWord = CheckBox_WholeWord.IsChecked == true;
                bool useWildcards = CheckBox_UseWildcards.IsChecked == true;
                bool includeProcs = CheckBox_StoredProcedures.IsChecked == true;
                bool includeViews = CheckBox_Views.IsChecked == true;
                bool includeFunctions = CheckBox_Functions.IsChecked == true;
                bool includeTables = CheckBox_Tables.IsChecked == true;

                DataTable results = await Task.Run(() => ExecuteSearch(searchText, allDatabases, wholeWord, useWildcards, includeProcs, includeViews, includeFunctions, includeTables));
                DataGrid_SearchResults.ItemsSource = results.DefaultView;
                TextBlock_ResultCount.Text = $"{results.Rows.Count} result(s)";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Search failed: {ex.Message}", "Quick Search");
                TextBlock_ResultCount.Text = "Search failed";
            }
            finally
            {
                Button_Search.IsEnabled = true;
            }
        }

        private DataTable ExecuteSearch(string searchText, bool allDatabases, bool wholeWord, bool useWildcards, bool includeProcs, bool includeViews, bool includeFunctions, bool includeTables)
        {
            List<string> databases = GetDatabasesToSearch(allDatabases);
            DataTable allResults = BuildResultTable();

            foreach (string dbName in databases)
            {
                DataTable rows = SearchDatabase(dbName, searchText, useWildcards, includeProcs, includeViews, includeFunctions, includeTables);
                foreach (DataRow row in rows.Rows)
                {
                    string sourceText = row["SourceText"]?.ToString() ?? string.Empty;
                    if (wholeWord && !Regex.IsMatch(sourceText, $@"\b{Regex.Escape(searchText)}\b", RegexOptions.IgnoreCase))
                    {
                        continue;
                    }

                    string preview = BuildPreview(sourceText, searchText, useWildcards);
                    allResults.Rows.Add(
                        row["DatabaseName"],
                        row["ObjectType"],
                        row["SchemaName"],
                        row["ObjectName"],
                        row["MatchLocation"],
                        preview);
                }
            }

            return allResults;
        }

        private List<string> GetDatabasesToSearch(bool allDatabases)
        {
            var list = new List<string>();

            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(selectedConnectionString)
            {
                InitialCatalog = "master"
            };

            using (var conn = new SqlConnection(builder.ConnectionString))
            {
                conn.Open();

                if (allDatabases)
                {
                    string sql = @"
SELECT [name]
FROM sys.databases
WHERE database_id > 4
  AND [state] = 0
  AND user_access = 0
ORDER BY [name];";

                    using (var cmd = new SqlCommand(sql, conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            list.Add(reader.GetString(0));
                        }
                    }
                }
                else
                {
                    list.Add(selectedDatabase);
                }
            }

            return list;
        }

        private DataTable SearchDatabase(string databaseName, string searchText, bool useWildcards, bool includeProcs, bool includeViews, bool includeFunctions, bool includeTables)
        {
            var result = new DataTable();

            string sql = $@"
USE [{databaseName}];

SELECT
    DB_NAME() AS DatabaseName,
    CASE
        WHEN o.[type] = 'P' THEN 'Stored Procedure'
        WHEN o.[type] = 'V' THEN 'View'
        ELSE 'Function'
    END AS ObjectType,
    s.[name] AS SchemaName,
    o.[name] AS ObjectName,
    'Definition' AS MatchLocation,
    m.[definition] AS SourceText
FROM sys.objects o
INNER JOIN sys.schemas s ON s.schema_id = o.schema_id
INNER JOIN sys.sql_modules m ON m.object_id = o.object_id
WHERE (
        (@includeProcs = 1 AND o.[type] = 'P') OR
        (@includeViews = 1 AND o.[type] = 'V') OR
        (@includeFunctions = 1 AND o.[type] IN ('FN', 'IF', 'TF'))
      )
  AND m.[definition] LIKE @pattern -- ESCAPE '\\'

UNION ALL

SELECT
    DB_NAME() AS DatabaseName,
    'Table' AS ObjectType,
    s.[name] AS SchemaName,
    t.[name] AS ObjectName,
    'Table Name' AS MatchLocation,
    t.[name] AS SourceText
FROM sys.tables t
INNER JOIN sys.schemas s ON s.schema_id = t.schema_id
WHERE @includeTables = 1
  AND t.[name] LIKE @pattern -- ESCAPE '\\'

UNION ALL

SELECT
    DB_NAME() AS DatabaseName,
    'Table' AS ObjectType,
    s.[name] AS SchemaName,
    t.[name] AS ObjectName,
    'Column' AS MatchLocation,
    c.[name] AS SourceText
FROM sys.tables t
INNER JOIN sys.schemas s ON s.schema_id = t.schema_id
INNER JOIN sys.columns c ON c.object_id = t.object_id
WHERE @includeTables = 1
  AND c.[name] LIKE @pattern -- ESCAPE '\\'

UNION ALL

SELECT
    DB_NAME() AS DatabaseName,
    CASE
        WHEN o.[type] = 'P' THEN 'Stored Procedure'
        WHEN o.[type] = 'V' THEN 'View'
        ELSE 'Function'
    END AS ObjectType,
    s.[name] AS SchemaName,
    o.[name] AS ObjectName,
    'Parameter' AS MatchLocation,
    p.[name] AS SourceText
FROM sys.parameters p
INNER JOIN sys.objects o ON o.object_id = p.object_id
INNER JOIN sys.schemas s ON s.schema_id = o.schema_id
WHERE p.parameter_id > 0
  AND (
        (@includeProcs = 1 AND o.[type] = 'P') OR
        (@includeViews = 1 AND o.[type] = 'V') OR
        (@includeFunctions = 1 AND o.[type] IN ('FN', 'IF', 'TF'))
      )
  AND p.[name] LIKE @pattern -- ESCAPE '\\';";

            string pattern = BuildPattern(searchText, useWildcards);

            using (var conn = new SqlConnection(selectedConnectionString))
            using (var cmd = new SqlCommand(sql, conn))
            using (var adapter = new SqlDataAdapter(cmd))
            {
                cmd.CommandTimeout = 120;
                cmd.Parameters.AddWithValue("@pattern", pattern);
                cmd.Parameters.AddWithValue("@includeProcs", includeProcs ? 1 : 0);
                cmd.Parameters.AddWithValue("@includeViews", includeViews ? 1 : 0);
                cmd.Parameters.AddWithValue("@includeFunctions", includeFunctions ? 1 : 0);
                cmd.Parameters.AddWithValue("@includeTables", includeTables ? 1 : 0);

                conn.Open();
                adapter.Fill(result);
            }

            return result;
        }

        private static string BuildPattern(string text, bool useWildcards)
        {
            if (useWildcards)
            {
                return text;
            }

            string escaped = text
                .Replace("\\", "\\\\")
                .Replace("%", "\\%")
                .Replace("_", "\\_")
                .Replace("[", "\\[");

            return $"%{escaped}%";
        }

        private static string BuildPreview(string sourceText, string searchText, bool useWildcards)
        {
            if (string.IsNullOrEmpty(sourceText))
            {
                return string.Empty;
            }

            if (!useWildcards)
            {
                int index = sourceText.IndexOf(searchText, StringComparison.OrdinalIgnoreCase);
                if (index >= 0)
                {
                    int start = Math.Max(0, index - 40);
                    int length = Math.Min(sourceText.Length - start, searchText.Length + 80);
                    string snippet = sourceText.Substring(start, length).Replace(Environment.NewLine, " ");
                    int snippetIndex = snippet.IndexOf(searchText, StringComparison.OrdinalIgnoreCase);
                    if (snippetIndex >= 0)
                    {
                        return snippet.Substring(0, snippetIndex) + "[" + snippet.Substring(snippetIndex, searchText.Length) + "]" + snippet.Substring(snippetIndex + searchText.Length);
                    }

                    return snippet;
                }
            }

            return sourceText.Length > 120
                ? sourceText.Substring(0, 120).Replace(Environment.NewLine, " ") + "..."
                : sourceText.Replace(Environment.NewLine, " ");
        }

        private bool AnyTypeSelected()
        {
            return CheckBox_StoredProcedures.IsChecked == true
                || CheckBox_Views.IsChecked == true
                || CheckBox_Functions.IsChecked == true
                || CheckBox_Tables.IsChecked == true;
        }

        private static DataTable BuildResultTable()
        {
            var table = new DataTable();
            table.Columns.Add("DatabaseName", typeof(string));
            table.Columns.Add("ObjectType", typeof(string));
            table.Columns.Add("SchemaName", typeof(string));
            table.Columns.Add("ObjectName", typeof(string));
            table.Columns.Add("MatchLocation", typeof(string));
            table.Columns.Add("MatchPreview", typeof(string));
            return table;
        }
    }
}