namespace AxialSqlTools
{
    using Microsoft.Data.SqlClient;
    using Newtonsoft.Json.Linq;
    using System;
    using System.Diagnostics;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.IO.Compression;
    using System.Net.Http;
    using System.Text;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Navigation;
    using static AxialSqlTools.AxialSqlToolsPackage;

    /// <summary>
    /// Interaction logic for SettingsWindowControl.
    /// </summary>
    public partial class SettingsWindowControl : UserControl
    {

        private string _queryHistoryConnectionString;

        private string tsqlFormatExample = @"while (1=0) 
begin 
select top 10
    c.CustomerID, getDate(),
    CASE WHEN o.TotalAmount > 1000 THEN 'High' ELSE 'Low' END AS OrderSize
FROM Customers c
JOIN Orders o ON c.CustomerID = o.CustomerID CROSS JOIN Regions r
WHERE c.IsActive = 1;

SELECT dbo.func(p.ProductID), p.ProductName FROM Products p; EXEC dbo.test @a = 0, @b = 1;
end
if 1=0 begin select 1; declare @a int, @b varchar(10) = ''
end
go
create procedure dbo.test @a int, @b int = 0
as select 1;
";
        /// <summary>
        /// Initializes a new instance of the <see cref="SettingsWindowControl"/> class.
        /// </summary>
        public SettingsWindowControl()
        {
            this.InitializeComponent();

            LoadSavedSettings();

            this.Loaded += UserControl_Loaded;

            SourceQueryPreview.Text = tsqlFormatExample;

            formatTSqlExample();

        }

        private void UserControl_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            LoadSavedSettings();
        }

        private void LoadSavedSettings()
        {
            try
            {

                ScriptFolder.Text = SettingsManager.GetTemplatesFolder();

                UseSnippets.IsChecked = SettingsManager.GetUseSnippets();
                SnippetFolder.Text = SettingsManager.GetSnippetFolder();

                _queryHistoryConnectionString = SettingsManager.GetQueryHistoryConnectionString();
                QueryHistoryTableName.Text = SettingsManager.GetQueryHistoryTableName();
                UpdateQueryHistoryConnectionDetails();

                RefreshQueryHistoryCreateScript();

                MyEmailAddress.Text = SettingsManager.GetMyEmail();

                SettingsManager.SmtpSettings smtpSettings = SettingsManager.GetSmtpSettings();

                SMTP_Server.Text = smtpSettings.ServerName;
                SMTP_Port.Text = smtpSettings.Port.ToString();
                SMTP_UserName.Text = smtpSettings.Username;
                SMTP_Password.Password = smtpSettings.Password;
                SMTP_EnableSSL.IsChecked = smtpSettings.EnableSsl;

                var tsqlCodeFormatSettings = SettingsManager.GetTSqlCodeFormatSettings();
                PreserveComments.IsChecked = tsqlCodeFormatSettings.preserveComments;
                RemoveNewLineAfterJoin.IsChecked = tsqlCodeFormatSettings.removeNewLineAfterJoin;
                AddTabAfterJoinOn.IsChecked = tsqlCodeFormatSettings.addTabAfterJoinOn;
                MoveCrossJoinToNewLine.IsChecked = tsqlCodeFormatSettings.moveCrossJoinToNewLine;
                FormatCaseAsMultiline.IsChecked = tsqlCodeFormatSettings.formatCaseAsMultiline;
                AddNewLineBetweenStatementsInBlocks.IsChecked = tsqlCodeFormatSettings.addNewLineBetweenStatementsInBlocks;
                BreakSprocParametersPerLine.IsChecked = tsqlCodeFormatSettings.breakSprocParametersPerLine;
                UppercaseBuiltInFunctions.IsChecked = tsqlCodeFormatSettings.uppercaseBuiltInFunctions;
                UnindentBeginEndBlocks.IsChecked = tsqlCodeFormatSettings.unindentBeginEndBlocks;
                BreakVariableDefinitionsPerLine.IsChecked = tsqlCodeFormatSettings.breakVariableDefinitionsPerLine;  
                BreakSprocDefinitionParametersPerLine.IsChecked = tsqlCodeFormatSettings.breakSprocDefinitionParametersPerLine;
                // BreakSelectFieldsAfterTopAndUnindent.IsChecked = tsqlCodeFormatSettings.breakSelectFieldsAfterTopAndUnindent;

                OpenAiApiKey.Password = SettingsManager.GetOpenAiApiKey();

                // Excel export settings
                var excelSettings = SettingsManager.GetExcelExportSettings();
                ExcelExportIncludeSourceQuery.IsChecked = excelSettings.includeSourceQuery;
                ExcelExportAddAutoFilter.IsChecked = excelSettings.addAutofilter;
                ExcelExportBoolsAsNumbers.IsChecked = excelSettings.exportBoolsAsNumbers;
                ExcelExportDefaultDirectory.Text = excelSettings.defaultDirectory;
                ExcelExportDefaultFilename.Text = excelSettings.defaultFileName;                              

            }
            catch (Exception ex)
            {
                _logger.Error(ex, "An exception occurred while loading settings");

                string msg = $"Error message: {ex.Message} \nInnerException: {ex.InnerException}";
                MessageBox.Show(msg, "Error");
            }

            try
            {
                GitHubToken.Password = WindowsCredentialHelper.LoadToken("AxialSqlTools_GitHubToken");
            }
            catch 
            {
                // ??
            }

        }

        private void UpdateQueryHistoryConnectionDetails()
        {

            if (string.IsNullOrWhiteSpace(_queryHistoryConnectionString))
            {
                Label_QueryHistoryConnectionInfo.Text = " < not configured > ";
            }
            else
            {
                try
                {

                    SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(_queryHistoryConnectionString);

                    string msg = string.Format("Server: {0}; Database: {1}; User ID: {2}", builder.DataSource, builder.InitialCatalog, builder.UserID);

                    Label_QueryHistoryConnectionInfo.Text = msg;

                }
                catch (Exception ex)
                {
                    Label_QueryHistoryConnectionInfo.Text = ex.Message;
                }
            }
        }

        private void Button_SaveScriptFolder_Click(object sender, RoutedEventArgs e)
        {
            SettingsManager.SaveTemplatesFolder(ScriptFolder.Text);

            SavedMessage();
        }


        private void Button_SaveSnippetFolder_Click(object sender, RoutedEventArgs e)
        {
            SettingsManager.SaveSnippetUse(UseSnippets.IsChecked.GetValueOrDefault(), SnippetFolder.Text);

            SavedMessage();
        }

        private void buttonDownloadAxialScripts_Click(object sender, RoutedEventArgs e)
        {
            string repoUrl = "https://github.com/Axial-SQL/AxialSqlTools/archive/main.zip";
            string targetFolderPath = "AxialSqlTools-main/query-library"; // Relative path inside the zip
            string targetPath = SettingsManager.GetTemplatesFolder();

            try
            {
                // Download the repo zip
                string tempZipPath = DownloadGitHubRepoZip(repoUrl);

                // Extract the specific folder from the zip
                ExtractSpecificFolderFromZip(tempZipPath, targetFolderPath, targetPath);

                MessageBox.Show("Axial SQL Tool Query Library has been downloaded", "Done");

            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format(System.Globalization.CultureInfo.CurrentUICulture, "An error occurred: '{0}'", ex.Message),
                    "Error");
            }

        }

        static string DownloadGitHubRepoZip(string url)
        {
            using (HttpClient client = new HttpClient())
            {              
                // Mimic a browser's User-Agent string
                client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3");
                client.DefaultRequestHeaders.Accept.ParseAdd("text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                client.DefaultRequestHeaders.AcceptLanguage.ParseAdd("en-US,en;q=0.5");

                string tempPath = Path.GetTempFileName() + ".zip";
                byte[] data = client.GetByteArrayAsync(url).GetAwaiter().GetResult();
                File.WriteAllBytes(tempPath, data);
                return tempPath;
            }
        }

        static void ExtractSpecificFolderFromZip(string zipPath, string folderPath, string destinationPath)
        {
            using (ZipArchive archive = ZipFile.OpenRead(zipPath))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.StartsWith(folderPath, StringComparison.OrdinalIgnoreCase))
                    {
                        string path = Path.Combine(destinationPath, entry.FullName.Substring(folderPath.Length + 1));

                        // Create subdirectory structure in destination, if needed
                        if (entry.FullName.EndsWith("/"))
                        {
                            Directory.CreateDirectory(path);
                        }
                        else
                        {
                            // Ensure directory exists
                            Directory.CreateDirectory(Path.GetDirectoryName(path));
                            // Check if file exists to avoid IOException
                            if (File.Exists(path))
                            {
                                File.Delete(path); // Delete the file if it exists.
                            }
                            entry.ExtractToFile(path, true);
                        }
                    }
                }
            }
            // Delete the temporary zip file after extraction
            File.Delete(zipPath);
        }

        private void SavedMessage()
        {
            MessageBox.Show(
                string.Format(System.Globalization.CultureInfo.CurrentUICulture, "The change has been saved", this.ToString()),
                "Setting saved");
        }

        private void buttonWikiPage_Click(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
        }

        private void ButtonSaveSmtpSettings_Click(object sender, RoutedEventArgs e)
        {
            
            SettingsManager.SmtpSettings smtpSettings = new SettingsManager.SmtpSettings()
            {
                ServerName = SMTP_Server.Text,
                Username = SMTP_UserName.Text,
                Password = SMTP_Password.Password,
                EnableSsl = SMTP_EnableSSL.IsChecked.GetValueOrDefault()
            };

            int smptPort = 587;
            bool success = int.TryParse(SMTP_Port.Text, out smptPort);
            smtpSettings.Port = smptPort;

            SettingsManager.SaveSmtpSettings(smtpSettings);

            SettingsManager.SaveMyEmail(MyEmailAddress.Text);

            SavedMessage();

        }

        private void Button_SaveApplyAdditionalFormat_Click(object sender, RoutedEventArgs e)
        {
            var settings = new SettingsManager.TSqlCodeFormatSettings
            {
                preserveComments = PreserveComments.IsChecked.GetValueOrDefault(false),
                removeNewLineAfterJoin = RemoveNewLineAfterJoin.IsChecked.GetValueOrDefault(false),
                addTabAfterJoinOn = AddTabAfterJoinOn.IsChecked.GetValueOrDefault(false),
                moveCrossJoinToNewLine = MoveCrossJoinToNewLine.IsChecked.GetValueOrDefault(false),
                formatCaseAsMultiline = FormatCaseAsMultiline.IsChecked.GetValueOrDefault(false),
                addNewLineBetweenStatementsInBlocks = AddNewLineBetweenStatementsInBlocks.IsChecked.GetValueOrDefault(false),
                breakSprocParametersPerLine = BreakSprocParametersPerLine.IsChecked.GetValueOrDefault(false),
                uppercaseBuiltInFunctions = UppercaseBuiltInFunctions.IsChecked.GetValueOrDefault(false),
                unindentBeginEndBlocks = UnindentBeginEndBlocks.IsChecked.GetValueOrDefault(false),
                breakVariableDefinitionsPerLine = BreakVariableDefinitionsPerLine.IsChecked.GetValueOrDefault(false),
                breakSprocDefinitionParametersPerLine = BreakSprocDefinitionParametersPerLine.IsChecked.GetValueOrDefault(false),
                // breakSelectFieldsAfterTopAndUnindent = BreakSelectFieldsAfterTopAndUnindent.IsChecked.GetValueOrDefault(false)
            };

            SettingsManager.SaveTSqlCodeFormatSettings(settings);
            SavedMessage();
        }

        private void button_SaveExcelExportSettings_Click(object sender, RoutedEventArgs e)
        {
            var settings = new SettingsManager.ExcelExportSettings
            {
                includeSourceQuery = ExcelExportIncludeSourceQuery.IsChecked.GetValueOrDefault(false),
                addAutofilter = ExcelExportAddAutoFilter.IsChecked.GetValueOrDefault(false),
                exportBoolsAsNumbers = ExcelExportBoolsAsNumbers.IsChecked.GetValueOrDefault(false),
                defaultDirectory = ExcelExportDefaultDirectory.Text,
                defaultFileName = ExcelExportDefaultFilename.Text
            };

            SettingsManager.SaveExcelExportSettings(settings);
            SavedMessage();
        }

        private void Hyperlink_RequestNavigateFormatQueryWiki(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
        }

        private void Button_SaveOpenAi_Click(object sender, RoutedEventArgs e)
        {
            SettingsManager.SaveOpenAiApiKey(OpenAiApiKey.Password);

            SavedMessage();
        }

        private void Button_SaveQueryHistory_Click(object sender, RoutedEventArgs e)
        {
            SettingsManager.SaveQueryHistoryConnectionString(_queryHistoryConnectionString);
            SettingsManager.SaveQueryHistoryTableName(QueryHistoryTableName.Text); 

            SavedMessage();

            RefreshQueryHistoryCreateScript();

        }

        private void Button_SelectDatabaseFromObjectExplorer_Click(object sender, RoutedEventArgs e)
        {

            var ci = ScriptFactoryAccess.GetCurrentConnectionInfo();

            _queryHistoryConnectionString = ci.FullConnectionString;

            UpdateQueryHistoryConnectionDetails();

        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Select templates folder";
                dialog.ShowNewFolderButton = true;

                // Show the dialog and check if the user selected a folder
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    // Set the selected folder path to the TextBox
                    ScriptFolder.Text = dialog.SelectedPath;
                }
            }
        }

        private void SnippetsBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Select snippets folder";
                dialog.ShowNewFolderButton = true;

                // Show the dialog and check if the user selected a folder
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    // Set the selected folder path to the TextBox
                    SnippetFolder.Text = dialog.SelectedPath;
                }
            }
        }

        private void Button_DisableQueryHistory_Click(object sender, RoutedEventArgs e)
        {

            _queryHistoryConnectionString = "";

            UpdateQueryHistoryConnectionDetails();

            RefreshQueryHistoryCreateScript();

        }

        private void formatTSqlExample()
        {
            var settings = new SettingsManager.TSqlCodeFormatSettings
            {
                preserveComments = PreserveComments.IsChecked.GetValueOrDefault(false),
                removeNewLineAfterJoin = RemoveNewLineAfterJoin.IsChecked.GetValueOrDefault(false),
                addTabAfterJoinOn = AddTabAfterJoinOn.IsChecked.GetValueOrDefault(false),
                moveCrossJoinToNewLine = MoveCrossJoinToNewLine.IsChecked.GetValueOrDefault(false),
                formatCaseAsMultiline = FormatCaseAsMultiline.IsChecked.GetValueOrDefault(false),
                addNewLineBetweenStatementsInBlocks = AddNewLineBetweenStatementsInBlocks.IsChecked.GetValueOrDefault(false),
                breakSprocParametersPerLine = BreakSprocParametersPerLine.IsChecked.GetValueOrDefault(false),
                uppercaseBuiltInFunctions = UppercaseBuiltInFunctions.IsChecked.GetValueOrDefault(false),
                unindentBeginEndBlocks = UnindentBeginEndBlocks.IsChecked.GetValueOrDefault(false),
                breakVariableDefinitionsPerLine = BreakVariableDefinitionsPerLine.IsChecked.GetValueOrDefault(false),
                breakSprocDefinitionParametersPerLine = BreakSprocDefinitionParametersPerLine.IsChecked.GetValueOrDefault(false),
                // breakSelectFieldsAfterTopAndUnindent = BreakSelectFieldsAfterTopAndUnindent.IsChecked.GetValueOrDefault(false)
            };

            FormattedQueryPreview.Text = TSqlFormatter.FormatCode(SourceQueryPreview.Text, settings);
        }

        private void formatSetting_Checked(object sender, RoutedEventArgs e)
        {
            formatTSqlExample();
        }

        private void formatSetting_Unchecked(object sender, RoutedEventArgs e)
        {
            formatTSqlExample();
        }

        private void buttonSaveGitHubSettings_Click(object sender, RoutedEventArgs e)
        {

            WindowsCredentialHelper.SaveToken("AxialSqlTools_GitHubToken", "AxialSqlTools_GitHubToken", GitHubToken.Password);

            SavedMessage();

        }


        private static string DefaultQueryHistoryTableName => "[dbo].[QueryHistory]";

        private void QueryHistoryTableName_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshQueryHistoryCreateScript();
        }

        private string EffectiveQueryHistoryTableName()
        {
            var name = QueryHistoryTableName?.Text;
            return string.IsNullOrWhiteSpace(name) ? DefaultQueryHistoryTableName : name.Trim();
        }

        private string GenerateQueryHistoryCreateTableScript(string tableName)
        {
            // Deterministic index names for display-only purposes
            string indexNameGuid = Guid.NewGuid().ToString();

            return $@"
IF OBJECT_ID(N'{tableName}', N'U') IS NULL
BEGIN
    CREATE TABLE {tableName} (
        [QueryID]           INT            IDENTITY (1, 1) NOT NULL,
        [StartTime]         DATETIME       NOT NULL,
        [FinishTime]        DATETIME       NOT NULL,
        [ElapsedTime]       VARCHAR (15)   NOT NULL,
        [TotalRowsReturned] BIGINT         NOT NULL,
        [ExecResult]        VARCHAR (100)  NOT NULL,
        [QueryText]         NVARCHAR (MAX) NOT NULL,
        [DataSource]        NVARCHAR (128) NOT NULL,
        [DatabaseName]      NVARCHAR (128) NOT NULL,
        [LoginName]         NVARCHAR (128) NOT NULL,
        [WorkstationId]     NVARCHAR (128) NOT NULL,
        PRIMARY KEY CLUSTERED ([QueryID]),
        INDEX [IDX_{indexNameGuid}_1] ([StartTime]),
        INDEX [IDX_{indexNameGuid}_2] ([FinishTime]),
        INDEX [IDX_{indexNameGuid}_3] ([DataSource]),
        INDEX [IDX_{indexNameGuid}_4] ([DatabaseName])
    );
    ALTER INDEX ALL ON {tableName} REBUILD WITH (DATA_COMPRESSION = PAGE);
END
".Trim();
        }

        private void RefreshQueryHistoryCreateScript()
        {
            try
            {
                QueryHistoryCreateScript.Text = GenerateQueryHistoryCreateTableScript(EffectiveQueryHistoryTableName());
            }
            catch (Exception ex)
            {
                QueryHistoryCreateScript.Text = $"-- Failed to generate script: {ex.Message}";
            }
        }



    }
}