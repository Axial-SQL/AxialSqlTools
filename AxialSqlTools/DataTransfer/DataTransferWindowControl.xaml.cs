namespace AxialSqlTools
{
    using Microsoft.VisualStudio.Shell;
    using System;
    using System.Data;
    using System.Data.SqlClient;
    using System.Diagnostics;
    using System.Diagnostics.CodeAnalysis;
    using System.Reflection;
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


    }
}