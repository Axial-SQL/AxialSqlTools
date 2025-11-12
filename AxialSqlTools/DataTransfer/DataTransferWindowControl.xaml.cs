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
    using System.Windows.Media;

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

            // Apply theme when control loads
            this.Loaded += (s, e) =>
            {
                ApplyThemeColors();
            };
            
            // Re-apply theme when control becomes visible (e.g., switching tabs or changing SSMS theme)
            this.IsVisibleChanged += (s, e) =>
            {
                if (this.IsVisible)
                {
                    ApplyThemeColors();
                }
            };

            // Re-apply theme when switching tabs within this page
            MainTabControl.SelectionChanged += (s, e) =>
            {
                if (this.IsVisible)
                {
                    // CRITICAL: Use Dispatcher to ensure visual tree is updated before theming
                    // When a tab is selected, WPF needs time to render its content into the visual tree
                    this.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (ThemeManager.IsDarkTheme())
                        {
                            var bgBrush = ThemeManager.GetBackgroundBrush();
                            var fgBrush = ThemeManager.GetForegroundBrush();
                            
                            // Theme the newly selected tab's content
                            if (MainTabControl.SelectedItem is TabItem selectedTab && selectedTab.Content is DependencyObject tabContent)
                            {
                                ApplyThemeToChildren(tabContent, bgBrush, fgBrush);
                            }

                            // Apply custom CheckBox style
                            ApplyCheckBoxStyle();
                        }
                    }), System.Windows.Threading.DispatcherPriority.Loaded);
                }
            };

            Button_CopyData.IsEnabled = false;
            Button_Cancel.Visibility = System.Windows.Visibility.Collapsed;

            CheckBox_TruncateTargetTableToPsql.IsChecked = true;
            CheckBox_CreateTargetTableToPsql.IsChecked = true;

            TextBox_TargetConnectionToPsql.Text = "Server=127.0.0.1;Port=5432;Database=postgres;User Id=postgres;Password=<password>;";

        }

        private void SqlBulkCopyOptionsExpander_Expanded(object sender, RoutedEventArgs e)
        {
            // When expander is expanded, its content is now in the visual tree
            // Apply theme after a brief delay to ensure visual tree is built
            this.Dispatcher.BeginInvoke(new Action(() =>
            {
                if (ThemeManager.IsDarkTheme() && SqlBulkCopyOptionsExpander.Content is DependencyObject expanderContent)
                {
                    var bgBrush = ThemeManager.GetBackgroundBrush();
                    var fgBrush = ThemeManager.GetForegroundBrush();
                    ApplyThemeToChildren(expanderContent, bgBrush, fgBrush);
                    
                    // Apply custom CheckBox style for checkboxes inside the expander
                    ApplyCheckBoxStyle();
                }
            }), System.Windows.Threading.DispatcherPriority.Loaded);
        }

        private bool _themeApplied = false;
        
        private void ApplyThemeColors()
        {
            try
            {
                // Always check current theme state - don't cache it
                bool isDark = ThemeManager.IsDarkTheme();
                
                if (!isDark)
                {
                    // Light mode - reset to default colors
                    this.ClearValue(Control.BackgroundProperty);
                    this.ClearValue(Control.ForegroundProperty);
                    // Clear theme from children too
                    ClearThemeFromChildren(this);
                    return;
                }

                // Dark mode - apply dark theme colors
                var bgBrush = ThemeManager.GetBackgroundBrush();
                var fgBrush = ThemeManager.GetForegroundBrush();

                this.Background = bgBrush;
                this.Foreground = fgBrush;

                // CRITICAL: Explicitly theme ALL tab items, not just visible ones
                if (MainTabControl != null)
                {
                    foreach (TabItem tabItem in MainTabControl.Items)
                    {
                        // Theme the tab header
                        tabItem.Foreground = fgBrush;
                        
                        // Theme the content of each tab (even if not visible)
                        if (tabItem.Content is DependencyObject tabContent)
                        {
                            ApplyThemeToChildren(tabContent, bgBrush, fgBrush);
                        }
                    }
                }

                // Recursively apply to all other children
                ApplyThemeToChildren(this, bgBrush, fgBrush);

                // Apply custom CheckBox style for dark mode
                ApplyCheckBoxStyle();
            }
            catch (Exception ex)
            {
                // Log but don't crash if theming fails
                System.Diagnostics.Debug.WriteLine($"Failed to apply theme: {ex.Message}");
            }
        }

        private void ClearThemeFromChildren(DependencyObject parent)
        {
            if (parent == null) return;

            int childCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);

                // Clear themed properties to restore defaults
                if (child is Label label)
                {
                    label.ClearValue(Label.ForegroundProperty);
                }
                else if (child is TextBlock textBlock)
                {
                    textBlock.ClearValue(TextBlock.ForegroundProperty);
                }
                else if (child is CheckBox checkBox)
                {
                    checkBox.ClearValue(CheckBox.ForegroundProperty);
                    checkBox.ClearValue(CheckBox.BorderBrushProperty);
                }
                else if (child is TextBox textBox)
                {
                    textBox.ClearValue(TextBox.BackgroundProperty);
                    textBox.ClearValue(TextBox.ForegroundProperty);
                }
                else if (child is RichTextBox richTextBox)
                {
                    richTextBox.ClearValue(RichTextBox.BackgroundProperty);
                    richTextBox.ClearValue(RichTextBox.ForegroundProperty);
                }
                else if (child is PasswordBox passwordBox)
                {
                    passwordBox.ClearValue(PasswordBox.BackgroundProperty);
                    passwordBox.ClearValue(PasswordBox.ForegroundProperty);
                }
                else if (child is GroupBox groupBox)
                {
                    groupBox.ClearValue(GroupBox.ForegroundProperty);
                }
                else if (child is Expander expander)
                {
                    expander.ClearValue(Expander.ForegroundProperty);
                }
                else if (child is TabControl tabControl)
                {
                    tabControl.ClearValue(TabControl.BackgroundProperty);
                    tabControl.ClearValue(TabControl.ForegroundProperty);
                }
                else if (child is TabItem tabItem)
                {
                    tabItem.ClearValue(TabItem.ForegroundProperty);
                }
                else if (child is Button button)
                {
                    button.ClearValue(Button.ForegroundProperty);
                }
                else if (child is Grid grid)
                {
                    grid.ClearValue(Grid.BackgroundProperty);
                }
                else if (child is StackPanel stackPanel)
                {
                    stackPanel.ClearValue(StackPanel.BackgroundProperty);
                }
                else if (child is Border border)
                {
                    border.ClearValue(Border.BackgroundProperty);
                }
                else if (child is DockPanel dockPanel)
                {
                    dockPanel.ClearValue(DockPanel.BackgroundProperty);
                }

                ClearThemeFromChildren(child);
            }
        }

        private void ApplyCheckBoxStyle()
        {
            // Use Dispatcher to ensure visual tree is fully constructed before applying style
            Dispatcher.BeginInvoke(new Action(() =>
            {
                var checkBoxStyle = this.TryFindResource("ThemedCheckBox") as Style;
                if (checkBoxStyle != null)
                {
                    ApplyCheckBoxStyleRecursive(this, checkBoxStyle);
                }
            }), System.Windows.Threading.DispatcherPriority.Loaded);
        }

        private void ApplyCheckBoxStyleRecursive(DependencyObject parent, Style style)
        {
            if (parent == null) return;

            int childCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);

                if (child is CheckBox checkBox)
                {
                    checkBox.Style = style;
                }

                ApplyCheckBoxStyleRecursive(child, style);
            }
        }

        private void ApplyThemeToChildren(DependencyObject parent, SolidColorBrush bgBrush, SolidColorBrush fgBrush)
        {
            if (parent == null) return;

            int childCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);

                // Apply foreground to all text-bearing controls - FORCE update even if already set
                if (child is Label label)
                {
                    label.Foreground = fgBrush;
                }
                else if (child is TextBlock textBlock)
                {
                    textBlock.Foreground = fgBrush;
                }
                else if (child is CheckBox checkBox)
                {
                    // Set foreground to white - this affects both the label text AND the checkmark glyph
                    checkBox.Foreground = fgBrush;
                    // Force the checkbox border to be visible
                    checkBox.BorderBrush = fgBrush;
                }
                else if (child is TextBox textBox)
                {
                    textBox.Background = bgBrush;
                    textBox.Foreground = fgBrush;
                }
                else if (child is RichTextBox richTextBox)
                {
                    richTextBox.Background = bgBrush;
                    richTextBox.Foreground = fgBrush;
                }
                else if (child is PasswordBox passwordBox)
                {
                    passwordBox.Background = bgBrush;
                    passwordBox.Foreground = fgBrush;
                }
                else if (child is GroupBox groupBox)
                {
                    groupBox.Foreground = fgBrush;
                }
                else if (child is Expander expander)
                {
                    expander.Foreground = fgBrush;
                }
                else if (child is TabControl tabControl)
                {
                    tabControl.Background = bgBrush;
                    tabControl.Foreground = fgBrush;
                }
                else if (child is TabItem tabItem)
                {
                    // CRITICAL: Set foreground on TabItem itself for header text
                    tabItem.Foreground = fgBrush;
                }
                else if (child is Button button)
                {
                    // CRITICAL: Set button foreground for text visibility
                    if (button.IsEnabled)
                    {
                        button.Foreground = fgBrush;
                        button.ClearValue(Button.BackgroundProperty); // Clear any disabled background
                        button.Opacity = 1.0; // Full opacity for enabled buttons
                    }
                    else
                    {
                        // Disabled buttons use opacity to look grayed out
                        button.Foreground = fgBrush; // Keep same color but use opacity
                        button.Opacity = 0.4; // Make it look disabled with transparency
                        button.Background = new SolidColorBrush(Color.FromRgb(60, 60, 60)); // Darker gray background
                    }
                }
                else if (child is Grid grid)
                {
                    // Theme grid backgrounds for tab content areas
                    grid.Background = bgBrush;
                }
                else if (child is Border border)
                {
                    // Theme borders too
                    if (border.Background != null)
                    {
                        border.Background = bgBrush;
                    }
                }

                // Recurse to theme nested controls
                ApplyThemeToChildren(child, bgBrush, fgBrush);
            }
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

                string postgresConnString = TextBox_TargetConnectionToPsql.Text;
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
        }
    }
}