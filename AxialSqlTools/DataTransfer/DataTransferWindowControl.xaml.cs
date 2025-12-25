namespace AxialSqlTools
{
    using Microsoft.VisualStudio.Shell;
    using MySqlConnector;
    using Npgsql;
    using System;
    using System.Data;
    using System.Data.SqlClient;
    using System.Diagnostics;
    using System.Diagnostics.CodeAnalysis;
    using System.Diagnostics.Tracing;
    using System.Reflection;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Documents;

    /// <summary>
    /// Interaction logic for DataTransferWindowControl.
    /// </summary>
    public partial class DataTransferWindowControl : UserControl
    {

        private CancellationTokenSource _cancellationTokenSource;
        private Stopwatch stopwatch;

        private string sourceConnectionString = "";
        private string targetConnectionString = "";

        private string sourceConnectionStringPsql = "";

        private string sourceConnectionStringMySql = "";
        private string sourceConnectionStringFromPsql = "";
        private string targetConnectionStringFromPsql = "";

        private string sourceConnectionStringFromMySql = "";
        private string targetConnectionStringFromMySql = "";

        private const int DefaultPostgresPort = 5432;
        private const int DefaultMySqlPort = 3306;


        static class SqlBulkCopyHelper
        {
            static FieldInfo rowsCopiedField = null;

            /// <summary>
            /// Gets the rows copied from the specified SqlBulkCopy object
            /// </summary>
            /// <param name="bulkCopy">The bulk copy.</param>
            /// <returns></returns>
            public static int GetRowsCopied(SqlBulkCopy bulkCopy)
            {
                if (rowsCopiedField == null)
                {
                    rowsCopiedField = typeof(SqlBulkCopy).GetField("_rowsCopied", BindingFlags.NonPublic | BindingFlags.GetField | BindingFlags.Instance);
                }

                return (int)rowsCopiedField.GetValue(bulkCopy);
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataTransferWindowControl"/> class.
        /// </summary>
        public DataTransferWindowControl()
        {
            this.InitializeComponent();

            Button_CopyData.IsEnabled = false;
            ButtonToPsql_CopyData.IsEnabled = false;
            ButtonToMySql_CopyData.IsEnabled = false;
            ButtonFromPsql_CopyData.IsEnabled = false;
            ButtonFromMySql_CopyData.IsEnabled = false;
            Button_Cancel.Visibility = System.Windows.Visibility.Collapsed;

            CheckBox_CreateTargetTableToPsql.IsChecked = true;
            CheckBox_CreateTargetTableToMySql.IsChecked = true;
            CheckBox_CreateTargetTableFromPsql.IsChecked = true;
            CheckBox_CreateTargetTableFromMySql.IsChecked = true;

            TextBox_TargetPsqlServer.Text = "127.0.0.1";
            TextBox_TargetPsqlPort.Text = DefaultPostgresPort.ToString();
            TextBox_TargetPsqlDatabase.Text = "postgres";
            TextBox_TargetPsqlUsername.Text = "postgres";
            PasswordBox_TargetPsqlPassword.Password = "postgres";

            TextBox_TargetMySqlServer.Text = "127.0.0.1";
            TextBox_TargetMySqlPort.Text = DefaultMySqlPort.ToString();
            TextBox_TargetMySqlDatabase.Text = "mysql";
            TextBox_TargetMySqlUsername.Text = "root";
            PasswordBox_TargetMySqlPassword.Password = "root";
            

            TextBox_SourcePsqlServer.Text = "127.0.0.1";
            TextBox_SourcePsqlPort.Text = DefaultPostgresPort.ToString();
            TextBox_SourcePsqlDatabase.Text = "postgres";
            TextBox_SourcePsqlUsername.Text = "postgres";
            PasswordBox_SourcePsqlPassword.Password = "postgres";

            TextBox_SourceMySqlServer.Text = "127.0.0.1";
            TextBox_SourceMySqlPort.Text = DefaultMySqlPort.ToString();
            TextBox_SourceMySqlDatabase.Text = "mysql";
            TextBox_SourceMySqlUsername.Text = "root";
            PasswordBox_SourceMySqlPassword.Password = "root";

        }

        private async void SqlToSql_CopyData_UpdateStatusAsync(object bulkCopySender, SqlRowsCopiedEventArgs eventArgs)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(_cancellationTokenSource.Token);

            TimeSpan ts = stopwatch.Elapsed;

            Label_CopyProgress.Content = $"Rows copied: {eventArgs.RowsCopied:#,0} in {(int)ts.TotalSeconds:#,0} sec.";

        }

        private void ButtonCopyData_Click(object sender, RoutedEventArgs e)
        {
            // await SqlToSql_CopyDataAsync();

            _cancellationTokenSource = new CancellationTokenSource();

            stopwatch = Stopwatch.StartNew();

            // update counter on the form with the last update time 
            ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {

                Button_CopyData.Visibility = System.Windows.Visibility.Collapsed;
                Button_Cancel.Visibility = System.Windows.Visibility.Visible;

                try
                {

                    int batchSize = 10000;
                    int totalRowsCopied = 0;

                    TextRange textRange = new TextRange(RichTextBox_SourceQuery.Document.ContentStart, RichTextBox_SourceQuery.Document.ContentEnd);
                    string sourceQuery = textRange.Text;

                    string targetTableName = TextBox_TargetTable.Text;

                    using (SqlConnection sourceConn = new SqlConnection(sourceConnectionString))
                    {
                        await sourceConn.OpenAsync();
                        using (SqlCommand cmd = new SqlCommand(sourceQuery, sourceConn))
                        {
                            cmd.CommandTimeout = 0;

                            Label_CopyProgress.Content = "Retrieving data from the source...";

                            using (SqlDataReader reader = await cmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess, _cancellationTokenSource.Token))
                            {

                                using (SqlConnection targetConn = new SqlConnection(targetConnectionString))
                                {
                                    await targetConn.OpenAsync();
                                    //var targetTransaction = targetConn.BeginTransaction();

                                    //-- Create target table
                                    string targetColumns = "";
                                    DataTable schemaTable = reader.GetSchemaTable();
                                    foreach (DataRow schemaRow in schemaTable.Rows)
                                    {
                                        string columnName = schemaRow[0].ToString();
                                        string sqlDataTypeName = GridAccess.GetColumnSqlType(schemaRow);

                                        targetColumns += (targetColumns == "" ? "" : ",\n");
                                        targetColumns += columnName + " " + sqlDataTypeName;
                                    }

                                    string targetTableCommand =
                                        $"IF OBJECT_ID('{targetTableName}') IS NULL \n" +
                                        $"CREATE TABLE {targetTableName} ({targetColumns})";

                                    using (SqlCommand targetCmd = new SqlCommand(targetTableCommand, targetConn)) //, targetTransaction))
                                    {
                                        await targetCmd.ExecuteNonQueryAsync();
                                    }

                                    SqlBulkCopyOptions options = SqlBulkCopyOptions.Default;

                                    // Combine options based on checkbox states
                                    if (KeepIdentityOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.KeepIdentity;

                                    if (CheckConstraintsOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.CheckConstraints;

                                    if (TableLockOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.TableLock;

                                    if (KeepNullsOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.KeepNulls;

                                    if (FireTriggersOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.FireTriggers;

                                    //if (UseInternalTransactionOption.IsChecked == true)
                                    //    options |= SqlBulkCopyOptions.UseInternalTransaction;

                                    if (AllowEncryptedValueModificationsOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.AllowEncryptedValueModifications;


                                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(targetConn, options, null))
                                    {

                                      
                                        bulkCopy.DestinationTableName = targetTableName;
                                        bulkCopy.BatchSize = batchSize;
                                        bulkCopy.NotifyAfter = batchSize;
                                        bulkCopy.SqlRowsCopied += SqlToSql_CopyData_UpdateStatusAsync;

                                        await bulkCopy.WriteToServerAsync(reader, _cancellationTokenSource.Token);

                                        totalRowsCopied = SqlBulkCopyHelper.GetRowsCopied(bulkCopy);
                                    }

                                    //targetTransaction.Commit();

                                }                               
                            }
                        }
                    }

                    TimeSpan ts = stopwatch.Elapsed;

                    Label_CopyProgress.Content = $"Completed | Total rows copied: {totalRowsCopied:#,0} in {(int)ts.TotalSeconds:#,0} sec.";

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Something went wrong: {ex.Message}", "DataTransferWindow");
                }

                Button_CopyData.Visibility = System.Windows.Visibility.Visible;
                Button_Cancel.Visibility = System.Windows.Visibility.Collapsed;

                stopwatch.Stop();

            });

        }

        private void ButtonCancel_Click(object sender, RoutedEventArgs e)
        {
            _cancellationTokenSource.Cancel();
        }

        private void Button_SelectSource_Click(object sender, RoutedEventArgs e)
        {
            var ci = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer();

            sourceConnectionString = ci.FullConnectionString;

            Label_SourceDescription.Content = $"Server: [{ci.ServerName}] / Database: [{ci.Database}]";

            SetCopyCommandAvailability();
        }

        private void Button_SelectSourceToMySql_Click(object sender, RoutedEventArgs e)
        {
            var ci = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer();

            sourceConnectionStringMySql = ci.FullConnectionString;

            Label_SourceDescriptionToMySql.Content = $"Server: [{ci.ServerName}] / Database: [{ci.Database}]";
            SetCopyCommandAvailabilityToMySql();
        }

        private void Button_SelectTargetFromPsql_Click(object sender, RoutedEventArgs e)
        {
            var ci = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer();

            targetConnectionStringFromPsql = ci.FullConnectionString;

            Label_TargetDescriptionFromPsql.Content = $"Server: [{ci.ServerName}] / Database: [{ci.Database}]";
            SetCopyCommandAvailabilityFromPsql();
        }

        private void Button_SelectTargetFromMySql_Click(object sender, RoutedEventArgs e)
        {
            var ci = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer();

            targetConnectionStringFromMySql = ci.FullConnectionString;

            Label_TargetDescriptionFromMySql.Content = $"Server: [{ci.ServerName}] / Database: [{ci.Database}]";
            SetCopyCommandAvailabilityFromMySql();
        }

        private void Button_SelectTarget_Click(object sender, RoutedEventArgs e)
        {
            var ci = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer();

            targetConnectionString = ci.FullConnectionString;

            Label_TargetDescription.Content = $"Server: [{ci.ServerName}] / Database: [{ci.Database}]";

            SetCopyCommandAvailability();
        }

        private void SetCopyCommandAvailability()
        {
            if (string.IsNullOrEmpty(sourceConnectionString) || string.IsNullOrEmpty(targetConnectionString))
            {
                Button_CopyData.IsEnabled = false;
            } else
            {
                Button_CopyData.IsEnabled = true;
            }
        }

        private void SetCopyCommandAvailabilityToPsql()
        {
            ButtonToPsql_CopyData.IsEnabled = !string.IsNullOrEmpty(sourceConnectionStringPsql);
        }

        private void SetCopyCommandAvailabilityToMySql()
        {
            ButtonToMySql_CopyData.IsEnabled = !string.IsNullOrEmpty(sourceConnectionStringMySql);
        }

        private void SetCopyCommandAvailabilityFromPsql()
        {
            ButtonFromPsql_CopyData.IsEnabled = !string.IsNullOrEmpty(targetConnectionStringFromPsql);
        }

        private void SetCopyCommandAvailabilityFromMySql()
        {
            ButtonFromMySql_CopyData.IsEnabled = !string.IsNullOrEmpty(targetConnectionStringFromMySql);
        }

        private int ParsePort(string portText, int defaultPort)
        {
            return int.TryParse(portText, out int parsedPort) ? parsedPort : defaultPort;
        }

        private string BuildPostgresConnectionString(string server, string portText, string database, string username, string password)
        {
            var builder = new NpgsqlConnectionStringBuilder
            {
                Host = server,
                Port = ParsePort(portText, DefaultPostgresPort),
                Database = database,
                Username = username,
                Password = password
            };

            return builder.ConnectionString;
        }

        private string BuildMySqlConnectionString(string server, string portText, string database, string username, string password)
        {
            uint parsedPort = (uint)ParsePort(portText, DefaultMySqlPort);
            var builder = new MySqlConnectionStringBuilder
            {
                Server = server,
                Port = parsedPort,
                Database = database,
                UserID = username,
                Password = password,
                AllowLoadLocalInfile = true
            };

            return builder.ConnectionString;
        }

        private string BuildTargetPostgresConnectionString()
        {
            return BuildPostgresConnectionString(
                TextBox_TargetPsqlServer.Text,
                TextBox_TargetPsqlPort.Text,
                TextBox_TargetPsqlDatabase.Text,
                TextBox_TargetPsqlUsername.Text,
                PasswordBox_TargetPsqlPassword.Password);
        }

        private string BuildSourcePostgresConnectionString()
        {
            return BuildPostgresConnectionString(
                TextBox_SourcePsqlServer.Text,
                TextBox_SourcePsqlPort.Text,
                TextBox_SourcePsqlDatabase.Text,
                TextBox_SourcePsqlUsername.Text,
                PasswordBox_SourcePsqlPassword.Password);
        }

        private string BuildTargetMySqlConnectionString()
        {
            return BuildMySqlConnectionString(
                TextBox_TargetMySqlServer.Text,
                TextBox_TargetMySqlPort.Text,
                TextBox_TargetMySqlDatabase.Text,
                TextBox_TargetMySqlUsername.Text,
                PasswordBox_TargetMySqlPassword.Password);
        }

        private string BuildSourceMySqlConnectionString()
        {
            return BuildMySqlConnectionString(
                TextBox_SourceMySqlServer.Text,
                TextBox_SourceMySqlPort.Text,
                TextBox_SourceMySqlDatabase.Text,
                TextBox_SourceMySqlUsername.Text,
                PasswordBox_SourceMySqlPassword.Password);
        }

        string MapSqlServerToPostgresType(string sqlServerType)
        {
            switch (sqlServerType)
            {
                case "System.Int32": return "INTEGER";
                case "System.Int64": return "BIGINT";
                case "System.Int16": return "SMALLINT";
                case "System.Decimal": return "NUMERIC";
                case "System.Double": return "DOUBLE PRECISION";
                case "System.Single": return "REAL";
                case "System.String": return "TEXT";
                case "System.Boolean": return "BOOLEAN";
                case "System.DateTime": return "TIMESTAMP";
                case "System.Guid": return "UUID";
                case "System.Byte[]": return "BYTEA";
                default: return "TEXT"; // Default to TEXT for unknown types
            }
        }

        string MapClrToMySqlType(Type clrType)
        {
            switch (Type.GetTypeCode(clrType))
            {
                case TypeCode.Int32: return "INT";
                case TypeCode.Int64: return "BIGINT";
                case TypeCode.Int16: return "SMALLINT";
                case TypeCode.Decimal: return "DECIMAL(38,10)";
                case TypeCode.Double: return "DOUBLE";
                case TypeCode.Single: return "FLOAT";
                case TypeCode.String: return "TEXT";
                case TypeCode.Boolean: return "BOOLEAN";
                case TypeCode.DateTime: return "DATETIME";
                default:
                    if (clrType == typeof(Guid)) return "CHAR(36)";
                    if (clrType == typeof(byte[])) return "LONGBLOB";
                    return "TEXT";
            }
        }

        string MapClrToSqlServerType(Type clrType)
        {
            switch (Type.GetTypeCode(clrType))
            {
                case TypeCode.Int32: return "INT";
                case TypeCode.Int64: return "BIGINT";
                case TypeCode.Int16: return "SMALLINT";
                case TypeCode.Decimal: return "DECIMAL(38,10)";
                case TypeCode.Double: return "FLOAT";
                case TypeCode.Single: return "REAL";
                case TypeCode.String: return "NVARCHAR(MAX)";
                case TypeCode.Boolean: return "BIT";
                case TypeCode.DateTime: return "DATETIME2";
                default:
                    if (clrType == typeof(Guid)) return "UNIQUEIDENTIFIER";
                    if (clrType == typeof(byte[])) return "VARBINARY(MAX)";
                    if (clrType == typeof(DateTimeOffset)) return "DATETIMEOFFSET";
                    return "NVARCHAR(MAX)";
            }
        }

        private async void ButtonToPsql_CopyData_Click(object sender, RoutedEventArgs e)
        {
            // Ensure a target table name is provided.
            if (string.IsNullOrEmpty(TextBox_TargetTableToPsql.Text))
            {
                TextBox_TargetTableToPsql.Text = $"data_export_{DateTime.Now:yyyyddMMHHmmss}";
            }

            // Create a cancellation token for the async operations.
            _cancellationTokenSource = new CancellationTokenSource();
            CancellationToken cancellationToken = _cancellationTokenSource.Token;

            stopwatch = Stopwatch.StartNew();

            try
            {
                // Update UI: hide the copy button and show the cancel button.
                ButtonToPsql_CopyData.Visibility = System.Windows.Visibility.Collapsed;
                ButtonToPsql_Cancel.Visibility = System.Windows.Visibility.Visible;

                string postgresConnString = BuildTargetPostgresConnectionString();
                // Get the SQL query from the rich text box.
                TextRange textRange = new TextRange(RichTextBox_SourceQueryToPsql.Document.ContentStart,
                                                    RichTextBox_SourceQueryToPsql.Document.ContentEnd);
                string sqlQuery = textRange.Text;
                string targetTable = TextBox_TargetTableToPsql.Text;

                // Open SQL Server connection and execute the query asynchronously.
                using (var sqlConn = new SqlConnection(sourceConnectionStringPsql))
                {
                    await sqlConn.OpenAsync(cancellationToken);
                    using (var sqlCmd = new SqlCommand(sqlQuery, sqlConn))
                    {
                        sqlCmd.CommandTimeout = 0;
                        // Use SequentialAccess for large data streams.
                        using (var reader = await sqlCmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess, cancellationToken))
                        {
                            // Get the schema information.
                            DataTable schemaTable = reader.GetSchemaTable();
                            if (schemaTable == null)
                                throw new Exception("Failed to retrieve schema from SQL Server.");

                            // Build the CREATE TABLE command for PostgreSQL.
                            StringBuilder createTableQuery = new StringBuilder($"CREATE TABLE IF NOT EXISTS {targetTable} (");
                            StringBuilder allTargetColumns = new StringBuilder();

                            foreach (DataRow row in schemaTable.Rows)
                            {
                                string columnName = row["ColumnName"].ToString();
                                string sqlServerType = row["DataType"].ToString();
                                string postgresType = MapSqlServerToPostgresType(sqlServerType);

                                createTableQuery.Append($"{columnName} {postgresType}, ");
                                allTargetColumns.Append($"{columnName},");
                            }

                            // Remove the trailing comma and space.
                            createTableQuery.Length -= 2;
                            createTableQuery.Append(");");

                            // Remove trailing comma from column list.
                            allTargetColumns.Length--;
                            string copyCommand = $"COPY {targetTable} ({allTargetColumns}) FROM STDIN (FORMAT BINARY)";

                            // Open PostgreSQL connection asynchronously.
                            using (var pgConn = new NpgsqlConnection(postgresConnString))
                            {
                                await pgConn.OpenAsync(cancellationToken);

                                // Optionally create the target table.
                                if (CheckBox_CreateTargetTableToPsql.IsChecked == true)
                                {
                                    using (var pgCmd = new NpgsqlCommand(createTableQuery.ToString(), pgConn))
                                    {
                                        await pgCmd.ExecuteNonQueryAsync(cancellationToken);
                                    }
                                }

                                // Optionally truncate the target table.
                                if (CheckBox_TruncateTargetTableToPsql.IsChecked == true)
                                {
                                    using (var pgCmd = new NpgsqlCommand($"TRUNCATE TABLE {targetTable} RESTART IDENTITY;", pgConn))
                                    {
                                        await pgCmd.ExecuteNonQueryAsync(cancellationToken);
                                    }
                                }

                                int rowCount = 0;
                                // Optionally perform the data copy.
                                if (CheckBox_SkipDataCopyToPsql.IsChecked == false)
                                {
                                    // Begin the PostgreSQL binary import asynchronously.
                                    using (var importer = await pgConn.BeginBinaryImportAsync(copyCommand, cancellationToken))
                                    {
                                        while (await reader.ReadAsync(cancellationToken))
                                        {
                                            // Start a new row in the COPY stream.
                                            importer.StartRow();

                                            for (int i = 0; i < reader.FieldCount; i++)
                                            {
                                                bool isNull = await reader.IsDBNullAsync(i, cancellationToken);
                                                var value = isNull ? DBNull.Value : reader.GetValue(i);
                                                // Write each column value asynchronously.
                                                await importer.WriteAsync(value, cancellationToken);
                                            }

                                            rowCount++;

                                            // Update progress every 1000 rows.
                                            if (rowCount % 1000 == 0)
                                            {
                                                TimeSpan elapsed = stopwatch.Elapsed;
                                                Label_CopyProgressToPsql.Content =
                                                    $"Rows copied: {rowCount:#,0} in {(int)elapsed.TotalSeconds:#,0} sec.";
                                            }

                                            // Check for cancellation.
                                            cancellationToken.ThrowIfCancellationRequested();
                                        }
                                        await importer.CompleteAsync(cancellationToken);
                                    }
                                }

                                // Final status update.
                                TimeSpan totalElapsed = stopwatch.Elapsed;
                                Label_CopyProgressToPsql.Content =
                                    $"Completed | Total rows copied: {rowCount:#,0} in {(int)totalElapsed.TotalSeconds:#,0} sec.";
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Data transfer has been cancelled.", "DataTransferWindow");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Something went wrong: {ex.Message}", "DataTransferWindow");
            }
            finally
            {
                // Restore the buttons and stop the stopwatch.
                ButtonToPsql_CopyData.Visibility = System.Windows.Visibility.Visible;
                ButtonToPsql_Cancel.Visibility = System.Windows.Visibility.Collapsed;
                stopwatch.Stop();
            }
        }


        private void ButtonToPsql_Cancel_Click(object sender, RoutedEventArgs e)
        {
            _cancellationTokenSource.Cancel();
        }

        private void Button_SelectSourceToPsql_Click(object sender, RoutedEventArgs e)
        {
            var ci = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer();

            sourceConnectionStringPsql = ci.FullConnectionString;

            Label_SourceDescriptionToPsql.Content = $"Server: [{ci.ServerName}] / Database: [{ci.Database}]";
            SetCopyCommandAvailabilityToPsql();
        }

        private async void ButtonToMySql_CopyData_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(TextBox_TargetTableToMySql.Text))
            {
                TextBox_TargetTableToMySql.Text = $"data_export_{DateTime.Now:yyyyddMMHHmmss}";
            }

            _cancellationTokenSource = new CancellationTokenSource();
            CancellationToken cancellationToken = _cancellationTokenSource.Token;

            stopwatch = Stopwatch.StartNew();

            ButtonToMySql_CopyData.Visibility = System.Windows.Visibility.Collapsed;
            ButtonToMySql_Cancel.Visibility = System.Windows.Visibility.Visible;

            try
            {
                TextRange textRange = new TextRange(RichTextBox_SourceQueryToMySql.Document.ContentStart,
                                                    RichTextBox_SourceQueryToMySql.Document.ContentEnd);
                string sqlQuery = textRange.Text;
                string targetTable = TextBox_TargetTableToMySql.Text;
                string targetMySqlConnString = BuildTargetMySqlConnectionString();

                using (var sqlConn = new SqlConnection(sourceConnectionStringMySql))
                {
                    await sqlConn.OpenAsync(cancellationToken);
                    using (var sqlCmd = new SqlCommand(sqlQuery, sqlConn))
                    {
                        sqlCmd.CommandTimeout = 0;
                        using (var reader = await sqlCmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess, cancellationToken))
                        {
                            DataTable schemaTable = reader.GetSchemaTable();
                            if (schemaTable == null)
                                throw new Exception("Failed to retrieve schema from SQL Server.");

                            StringBuilder createTableQuery = new StringBuilder($"CREATE TABLE IF NOT EXISTS `{targetTable}` (");
                            StringBuilder columnList = new StringBuilder();

                            foreach (DataRow row in schemaTable.Rows)
                            {
                                string columnName = row["ColumnName"].ToString();
                                var clrType = (Type)row["DataType"];
                                string mysqlType = MapClrToMySqlType(clrType);

                                createTableQuery.Append($"`{columnName}` {mysqlType}, ");
                                columnList.Append($"`{columnName}`,");
                            }

                            createTableQuery.Length -= 2;
                            createTableQuery.Append(");");
                            columnList.Length--;

                            using (var mySqlConn = new MySqlConnection(targetMySqlConnString))
                            {
                                await mySqlConn.OpenAsync(cancellationToken);

                                if (CheckBox_CreateTargetTableToMySql.IsChecked == true)
                                {
                                    using (var createCmd = new MySqlCommand(createTableQuery.ToString(), mySqlConn))
                                    {
                                        await createCmd.ExecuteNonQueryAsync(cancellationToken);
                                    }
                                }

                                if (CheckBox_TruncateTargetTableToMySql.IsChecked == true)
                                {
                                    using (var truncateCmd = new MySqlCommand($"TRUNCATE TABLE `{targetTable}`;", mySqlConn))
                                    {
                                        await truncateCmd.ExecuteNonQueryAsync(cancellationToken);
                                    }
                                }

                                if (CheckBox_SkipDataCopyToMySql.IsChecked == false)
                                {
                                    using (var localInfileCmd = new MySqlCommand("SET GLOBAL local_infile=1;", mySqlConn))
                                    {
                                        await localInfileCmd.ExecuteNonQueryAsync(cancellationToken);
                                    }

                                    using (var localInfileCheckCmd = new MySqlCommand("SHOW VARIABLES LIKE 'local_infile';", mySqlConn))
                                    using (var localInfileReader = await localInfileCheckCmd.ExecuteReaderAsync(cancellationToken))
                                    {
                                        if (await localInfileReader.ReadAsync(cancellationToken))
                                        {
                                            string localInfileValue = localInfileReader.GetString("Value");
                                            if (!string.Equals(localInfileValue, "ON", StringComparison.OrdinalIgnoreCase) &&
                                                !string.Equals(localInfileValue, "1", StringComparison.OrdinalIgnoreCase))
                                            {
                                                MessageBox.Show(
                                                    "MySQL local infile is disabled. Enable local_infile on the server (and ensure AllowLoadLocalInfile is true) to use bulk copy.",
                                                    "DataTransferWindow");
                                                return;
                                            }
                                        }
                                    }

                                    var bulkCopy = new MySqlBulkCopy(mySqlConn)
                                    {
                                        DestinationTableName = targetTable,
                                        BulkCopyTimeout = 0,
                                        NotifyAfter = 10000
                                    };

                                    bulkCopy.MySqlRowsCopied += async (s, args) =>
                                    {
                                        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
                                        TimeSpan elapsed = stopwatch.Elapsed;
                                        Label_CopyProgressToMySql.Content = $"Rows copied: {args.RowsCopied:#,0} in {(int)elapsed.TotalSeconds:#,0} sec.";
                                    };

                                    foreach (DataRow row in schemaTable.Rows)
                                    {
                                        string columnName = row["ColumnName"].ToString();
                                        int sourceOrdinal = Convert.ToInt32(row["ColumnOrdinal"]);
                                        bulkCopy.ColumnMappings.Add(new MySqlBulkCopyColumnMapping(sourceOrdinal, columnName));
                                    }

                                    var result = await bulkCopy.WriteToServerAsync(reader, cancellationToken);

                                    TimeSpan totalElapsed = stopwatch.Elapsed;
                                    Label_CopyProgressToMySql.Content =
                                        $"Completed | Total rows copied: {result.RowsInserted:#,0} in {(int)totalElapsed.TotalSeconds:#,0} sec.";
                                }
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Data transfer has been cancelled.", "DataTransferWindow");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Something went wrong: {ex.Message}", "DataTransferWindow");
            }
            finally
            {
                ButtonToMySql_CopyData.Visibility = System.Windows.Visibility.Visible;
                ButtonToMySql_Cancel.Visibility = System.Windows.Visibility.Collapsed;
                stopwatch.Stop();
            }
        }

        private void ButtonToMySql_Cancel_Click(object sender, RoutedEventArgs e)
        {
            _cancellationTokenSource.Cancel();
        }

        private async void ButtonFromPsql_CopyData_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(TextBox_TargetTableFromPsql.Text))
            {
                TextBox_TargetTableFromPsql.Text = $"data_import_{DateTime.Now:yyyyddMMHHmmss}";
            }

            _cancellationTokenSource = new CancellationTokenSource();
            CancellationToken cancellationToken = _cancellationTokenSource.Token;
            stopwatch = Stopwatch.StartNew();

            ButtonFromPsql_CopyData.Visibility = System.Windows.Visibility.Collapsed;
            ButtonFromPsql_Cancel.Visibility = System.Windows.Visibility.Visible;

            try
            {
                TextRange textRange = new TextRange(RichTextBox_SourceQueryFromPsql.Document.ContentStart,
                                                    RichTextBox_SourceQueryFromPsql.Document.ContentEnd);
                string sqlQuery = textRange.Text;
                string targetTable = TextBox_TargetTableFromPsql.Text;
                sourceConnectionStringFromPsql = BuildSourcePostgresConnectionString();

                using (var npgConn = new NpgsqlConnection(sourceConnectionStringFromPsql))
                {
                    await npgConn.OpenAsync(cancellationToken);
                    using (var cmd = new NpgsqlCommand(sqlQuery, npgConn))
                    {
                        cmd.CommandTimeout = 0;
                        using (var reader = await cmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess, cancellationToken))
                        {
                            DataTable schemaTable = reader.GetSchemaTable();
                            if (schemaTable == null)
                                throw new Exception("Failed to retrieve schema from PostgreSQL.");

                            string targetColumns = "";
                            foreach (DataRow schemaRow in schemaTable.Rows)
                            {
                                string columnName = schemaRow["ColumnName"].ToString();
                                var clrType = (Type)schemaRow["DataType"];
                                string sqlDataTypeName = MapClrToSqlServerType(clrType);

                                targetColumns += (targetColumns == "" ? "" : ",\n");
                                targetColumns += $"[{columnName}] {sqlDataTypeName}";
                            }

                            string targetTableCommand =
                                $"IF OBJECT_ID('{targetTable}') IS NULL \n" +
                                $"CREATE TABLE {targetTable} ({targetColumns})";

                            using (SqlConnection targetConn = new SqlConnection(targetConnectionStringFromPsql))
                            {
                                await targetConn.OpenAsync(cancellationToken);

                                if (CheckBox_CreateTargetTableFromPsql.IsChecked == true)
                                {
                                    using (SqlCommand targetCmd = new SqlCommand(targetTableCommand, targetConn))
                                    {
                                        await targetCmd.ExecuteNonQueryAsync(cancellationToken);
                                    }
                                }

                                if (CheckBox_TruncateTargetTableFromPsql.IsChecked == true)
                                {
                                    using (SqlCommand truncateCmd = new SqlCommand($"TRUNCATE TABLE {targetTable};", targetConn))
                                    {
                                        await truncateCmd.ExecuteNonQueryAsync(cancellationToken);
                                    }
                                }

                                if (CheckBox_SkipDataCopyFromPsql.IsChecked == false)
                                {
                                    SqlBulkCopyOptions options = SqlBulkCopyOptions.Default;

                                    if (KeepIdentityOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.KeepIdentity;

                                    if (CheckConstraintsOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.CheckConstraints;

                                    if (TableLockOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.TableLock;

                                    if (KeepNullsOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.KeepNulls;

                                    if (FireTriggersOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.FireTriggers;

                                    if (AllowEncryptedValueModificationsOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.AllowEncryptedValueModifications;

                                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(targetConn, options, null))
                                    {
                                        bulkCopy.DestinationTableName = targetTable;
                                        bulkCopy.BatchSize = 10000;
                                        bulkCopy.NotifyAfter = 10000;

                                        bulkCopy.SqlRowsCopied += async (s, args) =>
                                        {
                                            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
                                            TimeSpan elapsed = stopwatch.Elapsed;
                                            Label_CopyProgressFromPsql.Content = $"Rows copied: {args.RowsCopied:#,0} in {(int)elapsed.TotalSeconds:#,0} sec.";
                                        };

                                        foreach (DataRow schemaRow in schemaTable.Rows)
                                        {
                                            string columnName = schemaRow["ColumnName"].ToString();
                                            bulkCopy.ColumnMappings.Add(columnName, columnName);
                                        }

                                        await bulkCopy.WriteToServerAsync(reader, cancellationToken);

                                        int totalRowsCopied = SqlBulkCopyHelper.GetRowsCopied(bulkCopy);
                                        TimeSpan ts = stopwatch.Elapsed;
                                        Label_CopyProgressFromPsql.Content = $"Completed | Total rows copied: {totalRowsCopied:#,0} in {(int)ts.TotalSeconds:#,0} sec.";
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Data transfer has been cancelled.", "DataTransferWindow");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Something went wrong: {ex.Message}", "DataTransferWindow");
            }
            finally
            {
                ButtonFromPsql_CopyData.Visibility = System.Windows.Visibility.Visible;
                ButtonFromPsql_Cancel.Visibility = System.Windows.Visibility.Collapsed;
                stopwatch.Stop();
            }
        }

        private void ButtonFromPsql_Cancel_Click(object sender, RoutedEventArgs e)
        {
            _cancellationTokenSource.Cancel();
        }

        private async void ButtonFromMySql_CopyData_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(TextBox_TargetTableFromMySql.Text))
            {
                TextBox_TargetTableFromMySql.Text = $"data_import_{DateTime.Now:yyyyddMMHHmmss}";
            }

            _cancellationTokenSource = new CancellationTokenSource();
            CancellationToken cancellationToken = _cancellationTokenSource.Token;
            stopwatch = Stopwatch.StartNew();

            ButtonFromMySql_CopyData.Visibility = System.Windows.Visibility.Collapsed;
            ButtonFromMySql_Cancel.Visibility = System.Windows.Visibility.Visible;

            try
            {
                TextRange textRange = new TextRange(RichTextBox_SourceQueryFromMySql.Document.ContentStart,
                                                    RichTextBox_SourceQueryFromMySql.Document.ContentEnd);
                string sqlQuery = textRange.Text;
                string targetTable = TextBox_TargetTableFromMySql.Text;
                sourceConnectionStringFromMySql = BuildSourceMySqlConnectionString();

                using (var mySqlConn = new MySqlConnection(sourceConnectionStringFromMySql))
                {
                    await mySqlConn.OpenAsync(cancellationToken);
                    using (var cmd = new MySqlCommand(sqlQuery, mySqlConn))
                    {
                        cmd.CommandTimeout = 0;
                        using (var reader = await cmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess, cancellationToken))
                        {
                            DataTable schemaTable = reader.GetSchemaTable();
                            if (schemaTable == null)
                                throw new Exception("Failed to retrieve schema from MySQL.");

                            string targetColumns = "";
                            foreach (DataRow schemaRow in schemaTable.Rows)
                            {
                                string columnName = schemaRow["ColumnName"].ToString();
                                var clrType = (Type)schemaRow["DataType"];
                                string sqlDataTypeName = MapClrToSqlServerType(clrType);

                                targetColumns += (targetColumns == "" ? "" : ",\n");
                                targetColumns += $"[{columnName}] {sqlDataTypeName}";
                            }

                            string targetTableCommand =
                                $"IF OBJECT_ID('{targetTable}') IS NULL \n" +
                                $"CREATE TABLE {targetTable} ({targetColumns})";

                            using (SqlConnection targetConn = new SqlConnection(targetConnectionStringFromMySql))
                            {
                                await targetConn.OpenAsync(cancellationToken);

                                if (CheckBox_CreateTargetTableFromMySql.IsChecked == true)
                                {
                                    using (SqlCommand targetCmd = new SqlCommand(targetTableCommand, targetConn))
                                    {
                                        await targetCmd.ExecuteNonQueryAsync(cancellationToken);
                                    }
                                }

                                if (CheckBox_TruncateTargetTableFromMySql.IsChecked == true)
                                {
                                    using (SqlCommand truncateCmd = new SqlCommand($"TRUNCATE TABLE {targetTable};", targetConn))
                                    {
                                        await truncateCmd.ExecuteNonQueryAsync(cancellationToken);
                                    }
                                }

                                if (CheckBox_SkipDataCopyFromMySql.IsChecked == false)
                                {
                                    SqlBulkCopyOptions options = SqlBulkCopyOptions.Default;

                                    if (KeepIdentityOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.KeepIdentity;

                                    if (CheckConstraintsOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.CheckConstraints;

                                    if (TableLockOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.TableLock;

                                    if (KeepNullsOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.KeepNulls;

                                    if (FireTriggersOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.FireTriggers;

                                    if (AllowEncryptedValueModificationsOption.IsChecked == true)
                                        options |= SqlBulkCopyOptions.AllowEncryptedValueModifications;

                                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(targetConn, options, null))
                                    {
                                        bulkCopy.DestinationTableName = targetTable;
                                        bulkCopy.BatchSize = 10000;
                                        bulkCopy.NotifyAfter = 10000;

                                        bulkCopy.SqlRowsCopied += async (s, args) =>
                                        {
                                            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
                                            TimeSpan elapsed = stopwatch.Elapsed;
                                            Label_CopyProgressFromMySql.Content = $"Rows copied: {args.RowsCopied:#,0} in {(int)elapsed.TotalSeconds:#,0} sec.";
                                        };

                                        foreach (DataRow schemaRow in schemaTable.Rows)
                                        {
                                            string columnName = schemaRow["ColumnName"].ToString();
                                            bulkCopy.ColumnMappings.Add(columnName, columnName);
                                        }

                                        await bulkCopy.WriteToServerAsync(reader, cancellationToken);

                                        int totalRowsCopied = SqlBulkCopyHelper.GetRowsCopied(bulkCopy);
                                        TimeSpan ts = stopwatch.Elapsed;
                                        Label_CopyProgressFromMySql.Content = $"Completed | Total rows copied: {totalRowsCopied:#,0} in {(int)ts.TotalSeconds:#,0} sec.";
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Data transfer has been cancelled.", "DataTransferWindow");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Something went wrong: {ex.Message}", "DataTransferWindow");
            }
            finally
            {
                ButtonFromMySql_CopyData.Visibility = System.Windows.Visibility.Visible;
                ButtonFromMySql_Cancel.Visibility = System.Windows.Visibility.Collapsed;
                stopwatch.Stop();
            }
        }

        private void ButtonFromMySql_Cancel_Click(object sender, RoutedEventArgs e)
        {
            _cancellationTokenSource.Cancel();
        }
    }
}
