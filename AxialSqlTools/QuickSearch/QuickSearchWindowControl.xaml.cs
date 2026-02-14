using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Xml;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.UI.Design;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static AxialSqlTools.ScriptFactoryAccess;

namespace AxialSqlTools
{
    public partial class QuickSearchWindowControl : UserControl
    {
        private string selectedConnectionString;
        private string selectedDatabase;
        private string selectedServer;
        private CancellationTokenSource searchCancellationTokenSource;

        public QuickSearchWindowControl()
        {
            this.InitializeComponent();

            CheckBox_WholeWord.IsChecked = true;

            using (var stream = typeof(QuickSearchWindowControl).Assembly.GetManifestResourceStream("AxialSqlTools.QuickSearch.sql.xshd"))
            using (var reader = new XmlTextReader(stream))
            {
                SqlEditor.SyntaxHighlighting = HighlightingLoader.Load(reader, HighlightingManager.Instance);
            }

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
            SearchInputsGrid.IsEnabled = true;
        }

        private async void Button_Search_Click(object sender, RoutedEventArgs e)
        {
            await RunSearchAsync();
        }

        private async Task RunSearchAsync()
        {
            if (searchCancellationTokenSource != null)
            {
                searchCancellationTokenSource.Cancel();
                return;
            }

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

            searchCancellationTokenSource = new CancellationTokenSource();

            try
            {
                Button_Search.Content = "Cancel";
                DataGrid_SearchResults.ItemsSource = null;
                SqlEditor.Text = string.Empty;
                TextBlock_ResultCount.Text = "Searching...";

                bool allDatabases = CheckBox_AllDatabases.IsChecked == true;
                bool wholeWord = CheckBox_WholeWord.IsChecked == true;
                bool useWildcards = CheckBox_UseWildcards.IsChecked == true;
                bool includeProcs = CheckBox_StoredProcedures.IsChecked == true;
                bool includeViews = CheckBox_Views.IsChecked == true;
                bool includeFunctions = CheckBox_Functions.IsChecked == true;
                bool includeTables = CheckBox_Tables.IsChecked == true;
                bool includeAgentJobSteps = CheckBox_AgentJobSteps.IsChecked == true;

                CancellationToken cancellationToken = searchCancellationTokenSource.Token;
                var progress = new Progress<string>(databaseName =>
                {
                    TextBlock_ResultCount.Text = $"Searching [{databaseName}]...";
                });

                DataTable results = await Task.Run(() => ExecuteSearchAsync(searchText, allDatabases, wholeWord, useWildcards, includeProcs, includeViews, includeFunctions, includeTables, includeAgentJobSteps, progress, cancellationToken), cancellationToken);
                DataGrid_SearchResults.ItemsSource = results.DefaultView;
                TextBlock_ResultCount.Text = $"{results.Rows.Count} result(s)";
            }
            catch (OperationCanceledException)
            {
                TextBlock_ResultCount.Text = "Search canceled";
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Operation cancelled by user."))
                {
                    TextBlock_ResultCount.Text = "Search canceled";
                }
                else
                {
                    MessageBox.Show($"Search failed: {ex.Message}", "Quick Search");
                    TextBlock_ResultCount.Text = "Search failed";
                }           
            }
            finally
            {
                searchCancellationTokenSource?.Dispose();
                searchCancellationTokenSource = null;
                Button_Search.Content = "Search";
            }
        }

        private async void TextBox_SearchText_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter)
            {
                return;
            }

            e.Handled = true;
            await RunSearchAsync();
        }

        private async Task<DataTable> ExecuteSearchAsync(string searchText, bool allDatabases, bool wholeWord, bool useWildcards, bool includeProcs, bool includeViews, bool includeFunctions, bool includeTables, bool includeAgentJobSteps, IProgress<string> progress, CancellationToken cancellationToken)
        {
            List<string> databases = GetDatabasesToSearch(allDatabases);
            DataTable allResults = BuildResultTable();

            foreach (string dbName in databases)
            {
                cancellationToken.ThrowIfCancellationRequested();
                progress?.Report(dbName);

                DataTable rows;

                try
                {
                    rows = await SearchDatabaseAsync(
                        dbName,
                        searchText,
                        useWildcards,
                        includeProcs,
                        includeViews,
                        includeFunctions,
                        includeTables,
                        includeAgentJobSteps,
                        cancellationToken);
                }
                catch (SqlException ex)
                {
                    continue;
                }


                foreach (DataRow row in rows.Rows)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    string sourceText = row["SourceText"]?.ToString() ?? string.Empty;

                    string preview = BuildPreview(sourceText, searchText, useWildcards);
                    allResults.Rows.Add(
                        row["DatabaseName"],
                        row["ObjectType"],
                        row["SchemaName"],
                        row["ObjectName"],
                        row["MatchLocation"],
                        preview,
                        row["ScriptDatabaseName"],
                        row["ScriptSchemaName"],
                        row["ScriptObjectName"]);
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
                    WHERE [name] <> 'tempdb'
                      AND [state] = 0 --ONLINE
                      AND [user_access] = 0 --MULTI_USER
                      AND HAS_DBACCESS([name]) = 1
                    ORDER BY [name];

                    ";

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

        private async Task<DataTable> SearchDatabaseAsync(string databaseName, string searchText, bool useWildcards, bool includeProcs, bool includeViews, bool includeFunctions, bool includeTables, bool includeAgentJobSteps, CancellationToken cancellationToken)
        {

            DataTable result = BuildResultTable();

            string definitionSql = $@"
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
    m.[definition] AS SourceText,
    DB_NAME() AS ScriptDatabaseName,
    s.[name] AS ScriptSchemaName,
    o.[name] AS ScriptObjectName
FROM sys.objects o
INNER JOIN sys.schemas s ON s.schema_id = o.schema_id
INNER JOIN sys.sql_modules m ON m.object_id = o.object_id
WHERE (
        (@includeProcs = 1 AND o.[type] = 'P') OR
        (@includeViews = 1 AND o.[type] = 'V') OR
        (@includeFunctions = 1 AND o.[type] IN ('FN', 'IF', 'TF'))
      )

  AND m.[definition] LIKE @pattern ESCAPE '!'
  AND o.is_ms_shipped = 0 ;";

           string tableSql = $@"
SELECT
    DB_NAME() AS DatabaseName,
    'Table' AS ObjectType,
    s.[name] AS SchemaName,
    t.[name] AS ObjectName,
    'Table Name' AS MatchLocation,
    t.[name] AS SourceText,
    DB_NAME() AS ScriptDatabaseName,
    s.[name] AS ScriptSchemaName,
    t.[name] AS ScriptObjectName
FROM sys.tables t
INNER JOIN sys.schemas s ON s.schema_id = t.schema_id
WHERE @includeTables = 1
  AND t.[name] LIKE @pattern ESCAPE '!';";

            string columnSql = $@"
SELECT
    DB_NAME() AS DatabaseName,
    'Table' AS ObjectType,
    s.[name] AS SchemaName,
    t.[name] AS ObjectName,
    'Column' AS MatchLocation,
    c.[name] AS SourceText,
    DB_NAME() AS ScriptDatabaseName,
    s.[name] AS ScriptSchemaName,
    t.[name] AS ScriptObjectName
FROM sys.tables t
INNER JOIN sys.schemas s ON s.schema_id = t.schema_id
INNER JOIN sys.columns c ON c.object_id = t.object_id
WHERE @includeTables = 1
  AND c.[name] LIKE @pattern ESCAPE '!';";

            string parameterSql = $@"
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
    p.[name] AS SourceText,
    DB_NAME() AS ScriptDatabaseName,
    s.[name] AS ScriptSchemaName,
    o.[name] AS ScriptObjectName
FROM sys.parameters p
INNER JOIN sys.objects o ON o.object_id = p.object_id
INNER JOIN sys.schemas s ON s.schema_id = o.schema_id
WHERE p.parameter_id > 0
  AND (
        (@includeProcs = 1 AND o.[type] = 'P') OR
        (@includeViews = 1 AND o.[type] = 'V') OR
        (@includeFunctions = 1 AND o.[type] IN ('FN', 'IF', 'TF'))
      )
  AND p.[name] LIKE @pattern ESCAPE '!';";

            string agentJobsSql = @"
SELECT
    'msdb' AS DatabaseName,
    'SQL Agent Job Step' AS ObjectType,
    'dbo' AS SchemaName,
    j.[name] + N' / Step ' + CONVERT(varchar(12), js.step_id) + N' - ' + js.step_name AS ObjectName,
    'JobStep' AS MatchLocation,
    js.[command] AS SourceText,
    N'msdb' AS ScriptDatabaseName,
    N'dbo' AS ScriptSchemaName,
    j.[name] AS ScriptObjectName
FROM dbo.sysjobs j
INNER JOIN dbo.sysjobsteps js ON js.job_id = j.job_id
WHERE js.[command] LIKE @pattern ESCAPE '!'
   OR js.step_name LIKE @pattern ESCAPE '!'
   OR j.[name] LIKE @pattern;";

            string pattern = BuildPattern(searchText, useWildcards);

            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(selectedConnectionString)
            {
                InitialCatalog = databaseName
            };

            using (var conn = new SqlConnection(builder.ConnectionString))
            {
                await conn.OpenAsync(cancellationToken);

                await ExecuteSearchQueryAsync(conn, definitionSql, pattern, includeProcs, includeViews, includeFunctions, includeTables, result, cancellationToken);
                await ExecuteSearchQueryAsync(conn, tableSql, pattern, includeProcs, includeViews, includeFunctions, includeTables, result, cancellationToken);
                await ExecuteSearchQueryAsync(conn, columnSql, pattern, includeProcs, includeViews, includeFunctions, includeTables, result, cancellationToken);
                await ExecuteSearchQueryAsync(conn, parameterSql, pattern, includeProcs, includeViews, includeFunctions, includeTables, result, cancellationToken);

                
                if (databaseName == "msdb" && includeAgentJobSteps)
                    await ExecuteSearchQueryAsync(conn, agentJobsSql, pattern, includeProcs, includeViews, includeFunctions, includeTables, result, cancellationToken);

            }

            return result;
        }

        private static async Task ExecuteSearchQueryAsync(SqlConnection conn, string sql, string pattern, bool includeProcs, bool includeViews, bool includeFunctions, bool includeTables, DataTable aggregateResult, CancellationToken cancellationToken)
        {
            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.CommandTimeout = 120;
                cmd.Parameters.AddWithValue("@pattern", pattern);
                cmd.Parameters.AddWithValue("@includeProcs", includeProcs ? 1 : 0);
                cmd.Parameters.AddWithValue("@includeViews", includeViews ? 1 : 0);
                cmd.Parameters.AddWithValue("@includeFunctions", includeFunctions ? 1 : 0);
                cmd.Parameters.AddWithValue("@includeTables", includeTables ? 1 : 0);

                using (var reader = await cmd.ExecuteReaderAsync(cancellationToken))
                {
                    if (reader.HasRows)
                    {
                        var chunk = new DataTable();
                        chunk.Load(reader);
                        aggregateResult.Merge(chunk, true, MissingSchemaAction.Add);
                    }
                }
            }
        }
        private static string BuildPattern(string text, bool useWildcards)
        {

            string escaped = text.Replace("!", "!!");

            if (useWildcards)
            {
                return escaped;
            }

            escaped = escaped
                .Replace("%", "!%")
                .Replace("_", "!_")
                .Replace("[", "![");

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


        private void CheckBox_WholeWord_Checked(object sender, RoutedEventArgs e)
        {
            if (CheckBox_WholeWord.IsChecked == true)
            {
                CheckBox_UseWildcards.IsChecked = false;
            }
        }

        private void CheckBox_UseWildcards_Checked(object sender, RoutedEventArgs e)
        {
            if (CheckBox_UseWildcards.IsChecked == true)
            {
                CheckBox_WholeWord.IsChecked = false;
            }
        }

        private bool AnyTypeSelected()
        {
            return CheckBox_StoredProcedures.IsChecked == true
                || CheckBox_Views.IsChecked == true
                || CheckBox_Functions.IsChecked == true
                || CheckBox_Tables.IsChecked == true
                || CheckBox_AgentJobSteps.IsChecked == true;
        }

        private void Button_ScriptResult_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!(sender is Button button) || !(button.DataContext is DataRowView rowView))
            {
                return;
            }

            try
            { 
                string databaseName = rowView["ScriptDatabaseName"]?.ToString();
                string schemaName = rowView["ScriptSchemaName"]?.ToString();
                string objectName = rowView["ScriptObjectName"]?.ToString();

                string matchLocation = rowView["MatchLocation"]?.ToString();

                if (matchLocation == "JobStep")
                {
                    MessageBox.Show($"TODO - WIP", "WIP");
                    return;
                }

                string selectedObjectName = $"[{databaseName}].[{schemaName}].[{objectName}]";

                string fullScriptResult = ScriptObjectDefinition.GetText(AxialSqlToolsPackage.PackageInstance, selectedObjectName);
                SqlEditor.Text = fullScriptResult;

                var connectionInfo = ScriptFactoryAccess.GetCurrentConnectionInfo();

                ServiceCache.ScriptFactory.CreateNewBlankScript(ScriptType.Sql, connectionInfo.ActiveConnectionInfo, null);

                EnvDTE.TextDocument doc = (EnvDTE.TextDocument)ServiceCache.ExtensibilityModel.Application.ActiveDocument.Object(null);

                doc.EndPoint.CreateEditPoint().Insert(fullScriptResult);
            }
            catch(Exception ex)
            {
                MessageBox.Show($"Scripting failed: {ex.Message}", "Script Object");
            }           

        }

        private void DataGrid_SearchResults_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!(DataGrid_SearchResults.SelectedItem is DataRowView rowView))
            {
                return;
            }

            try
            {
                string matchLocation = rowView["MatchLocation"]?.ToString();
                if (matchLocation == "JobStep")
                {
                    SqlEditor.Text = rowView["MatchPreview"]?.ToString() ?? string.Empty;
                    return;
                }

                string databaseName = rowView["ScriptDatabaseName"]?.ToString();
                string schemaName = rowView["ScriptSchemaName"]?.ToString();
                string objectName = rowView["ScriptObjectName"]?.ToString();

                if (string.IsNullOrWhiteSpace(databaseName)
                    || string.IsNullOrWhiteSpace(schemaName)
                    || string.IsNullOrWhiteSpace(objectName))
                {
                    SqlEditor.Text = string.Empty;
                    return;
                }

                string selectedObjectName = $"[{databaseName}].[{schemaName}].[{objectName}]";
                SqlEditor.Text = ScriptObjectDefinition.GetText(AxialSqlToolsPackage.PackageInstance, selectedObjectName);
            }
            catch
            {
                SqlEditor.Text = string.Empty;
            }
        }

        private void ScriptObjectPlaceholder(string databaseName, string schemaName, string objectName)
        {
            MessageBox.Show($"TODO: Script object\nDatabase: {databaseName}\nSchema: {schemaName}\nObject: {objectName}", "Quick Search");
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
            table.Columns.Add("ScriptDatabaseName", typeof(string));
            table.Columns.Add("ScriptSchemaName", typeof(string));
            table.Columns.Add("ScriptObjectName", typeof(string));
            return table;
        }
    }
}
