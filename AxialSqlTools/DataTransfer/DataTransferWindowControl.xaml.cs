namespace AxialSqlTools
{
    using Microsoft.VisualStudio.Shell;
    using Npgsql;
    using NpgsqlTypes;
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
            Button_Cancel.Visibility = System.Windows.Visibility.Collapsed;

            CheckBox_TruncateTargetTableToPsql.IsChecked = true;

        }

        /// <summary>
        /// Handles click on the button by displaying a message box.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event args.</param>
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Justification = "Sample code")]
        [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1300:ElementMustBeginWithUpperCaseLetter", Justification = "Default event handler naming pattern")]
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            //MessageBox.Show(
            //    string.Format(System.Globalization.CultureInfo.CurrentUICulture, "Invoked '{0}'", this.ToString()),
            //    "DataTransferWindow");
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
                            using (SqlDataReader reader = await cmd.ExecuteReaderAsync())
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
            var ci = ScriptFactoryAccess.GetCurrentConnectionInfo();

            sourceConnectionString = ci.FullConnectionString;

            Label_SourceDescription.Content = $"Server: [{ci.ServerName}] / Database: [{ci.Database}]";

            SetCopyCommandAvailability();
        }

        private void Button_SelectTarget_Click(object sender, RoutedEventArgs e)
        {
            var ci = ScriptFactoryAccess.GetCurrentConnectionInfo();

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

        private void ButtonToPsql_CopyData_Click(object sender, RoutedEventArgs e)
        {

            stopwatch = Stopwatch.StartNew();

            try
            {                
                
                string postgresConnString = TextBox_TargetConnectionToPsql.Text;

                TextRange textRange = new TextRange(RichTextBox_SourceQueryToPsql.Document.ContentStart, RichTextBox_SourceQueryToPsql.Document.ContentEnd);
                string sqlQuery = textRange.Text;

                string targetTable = TextBox_TargetTableToPsql.Text;
                
                // Open SQL Server connection and execute the query.
                using (var sqlConn = new SqlConnection(sourceConnectionStringPsql))
                {
                    sqlConn.Open();
                    using (var sqlCmd = new SqlCommand(sqlQuery, sqlConn))
                    {
                        // Using CommandBehavior.SequentialAccess can help with large data streams.
                        using (var reader = sqlCmd.ExecuteReader(System.Data.CommandBehavior.SequentialAccess))
                        {

                            DataTable schemaTable = reader.GetSchemaTable();
                            if (schemaTable == null)
                            {
                                throw new Exception("Failed to retrieve schema from SQL Server.");
                            }

                            StringBuilder createTableQuery = new StringBuilder($"CREATE TABLE IF NOT EXISTS {targetTable} (");

                            string AllTargetColumns = "";

                            foreach (DataRow row in schemaTable.Rows)
                            {
                                string columnName = row["ColumnName"].ToString();
                                string sqlServerType = row["DataType"].ToString();
                                string postgresType = MapSqlServerToPostgresType(sqlServerType);

                                createTableQuery.Append($"{columnName} {postgresType}, ");

                                AllTargetColumns += columnName + ",";
                            }

                            createTableQuery.Length -= 2; // Remove last comma
                            createTableQuery.Append(");");

                            AllTargetColumns = AllTargetColumns.Substring(0, AllTargetColumns.Length - 1); // Remove last comma

                            // PostgreSQL COPY command. Make sure the columns match the source order.
                            string copyCommand = $"COPY {targetTable} ({AllTargetColumns}) FROM STDIN (FORMAT BINARY)";


                            // Open PostgreSQL connection.
                            using (var pgConn = new NpgsqlConnection(postgresConnString))
                            {
                                pgConn.Open();

                                using (var pgCmd = new NpgsqlCommand(createTableQuery.ToString(), pgConn))
                                {
                                    pgCmd.ExecuteNonQuery();
                                }

                                if (CheckBox_TruncateTargetTableToPsql.IsChecked == true)
                                {
                                    using (var pgCmd = new NpgsqlCommand($"TRUNCATE TABLE {targetTable} RESTART IDENTITY;", pgConn))
                                    {
                                        pgCmd.ExecuteNonQuery();
                                    }
                                }


                                // Begin the binary import into PostgreSQL.
                                using (var importer = pgConn.BeginBinaryImport(copyCommand))
                                {

                                    int j = 0;
                                    while (reader.Read())
                                    {
                                        // Start a new row in the PostgreSQL COPY stream.
                                        importer.StartRow();

                                        for (int i = 0; i < reader.FieldCount; i++)
                                        {
                                            var value = reader.IsDBNull(i) ? DBNull.Value : reader.GetValue(i);
                                            importer.Write(value);
                                        }

                                        j += 1;

                                        if (j % 1000 == 0)
                                        {
                                            TimeSpan tsj = stopwatch.Elapsed;

                                            Label_CopyProgressToPsql.Content = $"Rows copied: {j:#,0} in {(int)tsj.TotalSeconds:#,0} sec.";
                                        }

                                    }
                                    // Complete the import.
                                    importer.Complete();

                                    TimeSpan ts = stopwatch.Elapsed;

                                    Label_CopyProgressToPsql.Content = $"Completed | Total rows copied: {j:#,0} in {(int)ts.TotalSeconds:#,0} sec.";

                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Something went wrong: {ex.Message}", "DataTransferWindow");
            }    

            stopwatch.Stop();


        }

        private void ButtonToPsql_Cancel_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Button_SelectSourceToPsql_Click(object sender, RoutedEventArgs e)
        {
            var ci = ScriptFactoryAccess.GetCurrentConnectionInfo();

            sourceConnectionStringPsql = ci.FullConnectionString;

            Label_SourceDescriptionToPsql.Content = $"Server: [{ci.ServerName}] / Database: [{ci.Database}]";
        }
    }
}