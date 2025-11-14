namespace AxialSqlTools
{
    using Microsoft.VisualStudio.Shell;
    using Microsoft.Win32;
    using System;
    using System.Collections.Generic;
    using System.Data.SqlClient;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows;
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for the Data Import tool window.
    /// </summary>
    public partial class DataImportWindowControl : UserControl
    {
        private ScriptFactoryAccess.ConnectionInfo targetConnection;
        private string selectedExcelPath;
        private bool isImporting;

        public DataImportWindowControl()
        {
            InitializeComponent();
            UpdateStatus("Choose an Excel file to get started.");
        }

        private void ButtonBrowse_OnClick(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = false
            };

            bool? result = dialog.ShowDialog();
            if (result == true)
            {
                selectedExcelPath = dialog.FileName;
                TextBox_ExcelPath.Text = selectedExcelPath;

                if (string.IsNullOrWhiteSpace(TextBox_TargetTable.Text))
                {
                    TextBox_TargetTable.Text = Path.GetFileNameWithoutExtension(selectedExcelPath);
                }

                UpdateStatus($"Loaded {Path.GetFileName(selectedExcelPath)}. Select the target database.");
                UpdateImportButtonState();
            }
        }

        private void ButtonSelectTarget_OnClick(object sender, RoutedEventArgs e)
        {
            var connectionInfo = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer();
            if (connectionInfo == null)
            {
                MessageBox.Show("Please select a database in Object Explorer first.", "Data Import", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            targetConnection = connectionInfo;
            TextBox_Target.Text = targetConnection.DisplayName;

            UpdateStatus($"Target ready: {targetConnection.DisplayName}.");
            UpdateImportButtonState();
        }

        private void ButtonImport_OnClick(object sender, RoutedEventArgs e)
        {
            if (!EnsureSelections())
            {
                return;
            }

            string worksheet = string.IsNullOrWhiteSpace(TextBox_Worksheet.Text) ? null : TextBox_Worksheet.Text.Trim();
            bool firstRowHeaders = CheckBox_FirstRowHeaders.IsChecked == true;
            bool createTable = CheckBox_CreateTable.IsChecked == true;
            bool truncateTable = CheckBox_Truncate.IsChecked == true;
            string destinationTable = TextBox_TargetTable.Text.Trim();
            var connectionInfo = targetConnection;

            SetBusyState(true);

            ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {
                try
                {
                    await PerformImportAsync(worksheet, firstRowHeaders, createTable, truncateTable, destinationTable, connectionInfo);
                }
                catch (Exception ex)
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                    MessageBox.Show($"Import failed: {ex.Message}", "Data Import", MessageBoxButton.OK, MessageBoxImage.Error);
                    UpdateStatus("Import failed. Review the error and try again.");
                }
                finally
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                    SetBusyState(false);
                }
            });
        }

        private bool EnsureSelections()
        {
            if (string.IsNullOrWhiteSpace(selectedExcelPath))
            {
                MessageBox.Show("Select an Excel workbook first.", "Data Import", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (targetConnection == null)
            {
                MessageBox.Show("Select a target database from Object Explorer.", "Data Import", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (string.IsNullOrWhiteSpace(TextBox_TargetTable.Text))
            {
                MessageBox.Show("Provide a destination table name.", "Data Import", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            return true;
        }

        private void ButtonClear_OnClick(object sender, RoutedEventArgs e)
        {
            selectedExcelPath = null;
            targetConnection = null;
            TextBox_ExcelPath.Text = string.Empty;
            TextBox_Worksheet.Text = string.Empty;
            TextBox_Target.Text = string.Empty;
            TextBox_TargetTable.Text = string.Empty;
            CheckBox_CreateTable.IsChecked = true;
            CheckBox_Truncate.IsChecked = false;
            CheckBox_FirstRowHeaders.IsChecked = true;

            UpdateStatus("Choose an Excel file to get started.");
            UpdateImportButtonState();
        }

        private void TextBox_TargetTable_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateImportButtonState();
        }

        private void UpdateImportButtonState()
        {
            Button_Import.IsEnabled =
                !isImporting &&
                !string.IsNullOrWhiteSpace(selectedExcelPath) &&
                targetConnection != null &&
                !string.IsNullOrWhiteSpace(TextBox_TargetTable.Text);
        }

        private void UpdateStatus(string message)
        {
            TextBlock_Status.Text = message;
        }

        private void SetBusyState(bool importing)
        {
            isImporting = importing;

            Button_Browse.IsEnabled = !importing;
            Button_SelectTarget.IsEnabled = !importing;
            Button_Clear.IsEnabled = !importing;
            TextBox_Worksheet.IsEnabled = !importing;
            TextBox_TargetTable.IsEnabled = !importing;
            CheckBox_CreateTable.IsEnabled = !importing;
            CheckBox_Truncate.IsEnabled = !importing;
            CheckBox_FirstRowHeaders.IsEnabled = !importing;

            UpdateImportButtonState();
        }

        private async Task PerformImportAsync(string worksheet, bool firstRowHeaders, bool createTable, bool truncateTable, string destinationTable, ScriptFactoryAccess.ConnectionInfo connectionInfo)
        {
            await UpdateStatusAsync("Reading Excel file...");

            ExcelImport.WorksheetData worksheetData = await Task.Run(() =>
                ExcelImport.ReadWorksheet(selectedExcelPath, worksheet, firstRowHeaders));

            await UpdateStatusAsync($"Loaded '{worksheetData.WorksheetName}' with {worksheetData.Table.Rows.Count:#,0} rows. Preparing destination table...");

            await ImportIntoSqlAsync(worksheetData, destinationTable, createTable, truncateTable, connectionInfo);

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            UpdateStatus($"Imported {worksheetData.Table.Rows.Count:#,0} rows into {destinationTable} on {connectionInfo.DisplayName}.");
            MessageBox.Show(
                $"Successfully imported {worksheetData.Table.Rows.Count:#,0} rows from {Path.GetFileName(selectedExcelPath)} into {connectionInfo.DisplayName} ({destinationTable}).",
                "Data Import",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private async Task ImportIntoSqlAsync(ExcelImport.WorksheetData worksheetData, string destinationTable, bool createTable, bool truncateTable, ScriptFactoryAccess.ConnectionInfo connectionInfo)
        {
            string quotedTableName = SqlIdentifierHelper.QuoteQualifiedName(destinationTable);
            string tableLiteral = EscapeForSqlLiteral(quotedTableName);

            using (SqlConnection connection = new SqlConnection(connectionInfo.FullConnectionString))
            {
                await connection.OpenAsync();

                await UpdateStatusAsync("Ensuring destination table exists...");

                if (createTable)
                {
                    string createScript = BuildCreateTableScript(tableLiteral, quotedTableName, worksheetData.Columns);
                    using (SqlCommand command = new SqlCommand(createScript, connection))
                    {
                        await command.ExecuteNonQueryAsync();
                    }
                }
                else
                {
                    string existsScript =
                        $"IF OBJECT_ID(N'{tableLiteral}', 'U') IS NULL BEGIN THROW 50000, 'Destination table was not found.', 1; END";

                    using (SqlCommand command = new SqlCommand(existsScript, connection))
                    {
                        await command.ExecuteNonQueryAsync();
                    }
                }

                if (truncateTable)
                {
                    await UpdateStatusAsync("Truncating destination table...");
                    using (SqlCommand command = new SqlCommand($"TRUNCATE TABLE {quotedTableName};", connection))
                    {
                        await command.ExecuteNonQueryAsync();
                    }
                }

                await UpdateStatusAsync("Copying rows into SQL Server...");

                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection, SqlBulkCopyOptions.TableLock, null))
                {
                    bulkCopy.DestinationTableName = quotedTableName;
                    bulkCopy.BatchSize = 5000;
                    bulkCopy.BulkCopyTimeout = 0;

                    foreach (ExcelImport.ExcelColumnMetadata column in worksheetData.Columns)
                    {
                        bulkCopy.ColumnMappings.Add(column.Name, column.Name);
                    }

                    await bulkCopy.WriteToServerAsync(worksheetData.Table);
                }
            }
        }

        private static string BuildCreateTableScript(string tableLiteral, string quotedTableName, IReadOnlyList<ExcelImport.ExcelColumnMetadata> columns)
        {
            string columnDefinitions = string.Join(
                ",\n        ",
                columns.Select(column => $"{SqlIdentifierHelper.QuoteIdentifier(column.Name)} {column.SqlType} NULL"));

            return
                $"IF OBJECT_ID(N'{tableLiteral}', 'U') IS NULL\n" +
                "BEGIN\n" +
                $"    CREATE TABLE {quotedTableName} (\n" +
                $"        {columnDefinitions}\n" +
                "    );\n" +
                "END";
        }

        private static string EscapeForSqlLiteral(string value)
        {
            return string.IsNullOrEmpty(value) ? string.Empty : value.Replace("'", "''");
        }

        private async Task UpdateStatusAsync(string message)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            UpdateStatus(message);
        }

        private static class SqlIdentifierHelper
        {
            public static string QuoteQualifiedName(string input)
            {
                var parts = GetNameParts(input);
                if (parts.Count == 0)
                {
                    throw new InvalidOperationException("Destination table name is invalid.");
                }

                return string.Join(".", parts.Select(QuoteIdentifier));
            }

            public static string QuoteIdentifier(string identifier)
            {
                if (string.IsNullOrWhiteSpace(identifier))
                {
                    throw new InvalidOperationException("Column names cannot be empty.");
                }

                string sanitized = identifier.Replace("]", "]]" );
                return $"[{sanitized}]";
            }

            private static List<string> GetNameParts(string input)
            {
                var parts = new List<string>();
                if (string.IsNullOrWhiteSpace(input))
                {
                    return parts;
                }

                StringBuilder current = new StringBuilder();
                bool insideBrackets = false;

                foreach (char ch in input)
                {
                    if (ch == '[')
                    {
                        insideBrackets = true;
                        continue;
                    }

                    if (ch == ']')
                    {
                        insideBrackets = false;
                        continue;
                    }

                    if (ch == '.' && !insideBrackets)
                    {
                        AddPart(current, parts);
                        continue;
                    }

                    current.Append(ch);
                }

                AddPart(current, parts);

                if (parts.Count == 0)
                {
                    parts.Add(input.Trim());
                }

                return parts;
            }

            private static void AddPart(StringBuilder builder, List<string> parts)
            {
                if (builder.Length == 0)
                {
                    return;
                }

                string value = builder.ToString().Trim();
                builder.Clear();

                if (!string.IsNullOrWhiteSpace(value))
                {
                    parts.Add(value);
                }
            }
        }
    }
}
