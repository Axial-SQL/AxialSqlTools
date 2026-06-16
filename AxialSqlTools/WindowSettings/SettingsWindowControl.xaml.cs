namespace AxialSqlTools
{
    using Microsoft.Data.SqlClient;
    using Microsoft.SqlServer.Management.UI.VSIntegration;
    using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;
    using Microsoft.SqlServer.TransactSql.ScriptDom;
    using Newtonsoft.Json.Linq;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.IO.Compression;
    using System.Net.Http;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Windows;
    using System.Windows.Controls;
    using System.Collections.ObjectModel;
    using System.Windows.Media;
    using System.Windows.Navigation;
    using Microsoft.VisualBasic;
    using static AxialSqlTools.AxialSqlToolsPackage;

    /// <summary>
    /// Interaction logic for SettingsWindowControl.
    /// </summary>
    public partial class SettingsWindowControl : UserControl
    {
        private const string QueryHistoryStorageModeDatabase = "Database";
        private const string QueryHistoryStorageModeTextFiles = "TextFiles";
        private const string QueryHistoryStorageModeDisabled = "Disabled";

        private string _queryHistoryConnectionString;
        private readonly ToolWindowThemeController _themeController;
        private bool updateResultSubscribed;
        private ObservableCollection<SettingsManager.ConnectionColorRule> _connectionColorRules;

        private string tsqlFormatExample = @"-- comment
IF 0 = 1
BEGIN
    DECLARE @a int = 1, @b varchar(10) = 'x';
    DECLARE @t TABLE (id int, label varchar(10));
    DECLARE @u TABLE (id int);

    WITH c AS (
        SELECT TOP 2 t.id, GETDATE() AS dt, CASE WHEN t.id > @a THEN 'big' ELSE 'small' END AS label
        FROM @t AS t JOIN @u AS u ON u.id = t.id CROSS JOIN @u AS x
    )
    SELECT c.id, c.dt, c.label -- inline comment
    FROM c
    WHERE c.label <> @b;

    EXEC dbo.p @a = @a, @b = @b;
END
";
        /// <summary>
        /// Initializes a new instance of the <see cref="SettingsWindowControl"/> class.
        /// </summary>
        public SettingsWindowControl()
        {
            this.InitializeComponent();

            _connectionColorRules = new ObservableCollection<SettingsManager.ConnectionColorRule>();
            ConnectionColorRulesListView.ItemsSource = _connectionColorRules;

            _themeController = new ToolWindowThemeController(this, ApplyThemeBrushResources);

            this.Loaded += UserControl_Loaded;
            this.Unloaded += UserControl_Unloaded;

            SourceQueryPreview.Text = tsqlFormatExample;
            FreehandSourceEditor.Text = tsqlFormatExample;
            FreehandFormatEditor.Text = tsqlFormatExample;

            formatTSqlExample();

        }

        private void UserControl_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            SubscribeToUpdateResultChanges();
            LoadSavedSettings();
        }

        private void UserControl_Unloaded(object sender, System.Windows.RoutedEventArgs e)
        {
            UnsubscribeFromUpdateResultChanges();
        }

        private void SubscribeToUpdateResultChanges()
        {
            if (updateResultSubscribed)
            {
                return;
            }

            UpdateChecker.LastUpdateResultChanged += UpdateChecker_LastUpdateResultChanged;
            updateResultSubscribed = true;
        }

        private void UnsubscribeFromUpdateResultChanges()
        {
            if (!updateResultSubscribed)
            {
                return;
            }

            UpdateChecker.LastUpdateResultChanged -= UpdateChecker_LastUpdateResultChanged;
            updateResultSubscribed = false;
        }

        private void ApplyThemeBrushResources()
        {
            ToolWindowThemeResources.ApplySharedTheme(this);

            ApplyGoogleSheetsAuthorizationBrush();
        }

        private Brush GetThemedStatusBrush(bool isSuccess)
        {
            string key = isSuccess ? "AxialThemeStatusSuccessBrush" : "AxialThemeStatusErrorBrush";
            return Resources[key] as Brush
                ?? (isSuccess ? new SolidColorBrush(Color.FromRgb(0x10, 0x7C, 0x10)) : new SolidColorBrush(Color.FromRgb(0xA1, 0x26, 0x0D)));
        }

        private void ApplyGoogleSheetsAuthorizationBrush()
        {
            if (GoogleSheetsRefreshTokenLabel == null)
            {
                return;
            }

            bool isAuthorized = string.Equals(GoogleSheetsRefreshTokenLabel.Text, "Authorized", StringComparison.OrdinalIgnoreCase);
            GoogleSheetsRefreshTokenLabel.Foreground = GetThemedStatusBrush(isAuthorized);
        }

        private void LoadSavedSettings()
        {
            try
            {

                ScriptFolder.Text = SettingsManager.GetTemplatesFolder();

                var snippetSettings = SettingsManager.GetSnippetSettings();
                UseSnippets.IsChecked = snippetSettings.useSnippets;
                SnippetFolder.Text = snippetSettings.snippetFolder;
                SnippetReplaceMode.SelectedValue = snippetSettings.replaceKey.ToString();

                _queryHistoryConnectionString = SettingsManager.GetQueryHistoryConnectionString();
                QueryHistoryTableName.Text = SettingsManager.GetQueryHistoryTableName();
                QueryHistoryTextFilesInfo.Text = SettingsManager.GetQueryHistoryTextFileFolder();
                SelectQueryHistoryStorageType(SettingsManager.GetQueryHistoryStorageMode());
                UpdateQueryHistoryStorageControls();
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
                ApplyTSqlCodeFormatSettingsToUi(tsqlCodeFormatSettings);
                FormatModeFreehand.IsChecked = tsqlCodeFormatSettings.useFreehandFormatMode;
                FormatModeOptions.IsChecked = !tsqlCodeFormatSettings.useFreehandFormatMode;
                FreehandSourceEditor.Text = tsqlFormatExample;
                FreehandFormatEditor.Text = TSqlFormatter.FormatCode(tsqlFormatExample, tsqlCodeFormatSettings);
                UpdateFormatModeUi();

                OpenAiApiKey.Password = SettingsManager.GetOpenAiApiKey();

                // Excel export settings
                var excelSettings = SettingsManager.GetExcelExportSettings();
                ExcelExportIncludeSourceQuery.IsChecked = excelSettings.includeSourceQuery;
                ExcelExportAddAutoFilter.IsChecked = excelSettings.addAutofilter;
                ExcelExportBoolsAsNumbers.IsChecked = excelSettings.exportBoolsAsNumbers;
                ExcelExportDefaultDirectory.Text = excelSettings.defaultDirectory;
                ExcelExportDefaultFilename.Text = excelSettings.defaultFileName;

                var googleSettings = SettingsManager.GetGoogleSheetsSettings();
                GoogleSheetsIncludeSourceQuery.IsChecked = googleSettings.includeSourceQuery;
                GoogleSheetsExportBoolsAsNumbers.IsChecked = googleSettings.exportBoolsAsNumbers;
                GoogleSheetsDefaultSpreadsheetName.Text = googleSettings.defaultSpreadsheetName;
                GoogleSheetsClientId.Text = googleSettings.clientId;
                GoogleSheetsClientSecret.Password = googleSettings.clientSecret;
                UpdateGoogleSheetsStatus(googleSettings.refreshToken);

                EnableUpdateChecks.IsChecked = SettingsManager.GetEnableUpdateChecks();
                UpdateUpdateStatus();

                LoadConnectionColorRules();

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
            var snippetSettings = SettingsManager.GetSnippetSettings();
            snippetSettings.useSnippets = UseSnippets.IsChecked.GetValueOrDefault();
            snippetSettings.snippetFolder = SnippetFolder.Text;
            snippetSettings.replaceKey = GetSelectedSnippetReplaceKey();

            SettingsManager.SaveSnippetSettings(snippetSettings);

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

        private SettingsManager.SnippetReplaceKey GetSelectedSnippetReplaceKey()
        {
            var selectedValue = SnippetReplaceMode.SelectedValue as string;
            if (Enum.TryParse(selectedValue, out SettingsManager.SnippetReplaceKey key))
            {
                return key;
            }

            return SettingsManager.SnippetReplaceKey.Enter;
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
            if (FormatModeFreehand.IsChecked == true && !ValidateFreehandFormatSample())
            {
                return;
            }

            var settings = FormatModeFreehand.IsChecked == true
                ? LearnTSqlFormatSettingsFromSample(FreehandFormatEditor.Text)
                : BuildTSqlCodeFormatSettingsFromUi();

            settings.useFreehandFormatMode = FormatModeFreehand.IsChecked == true;

            SettingsManager.SaveTSqlCodeFormatSettings(settings);
            ApplyTSqlCodeFormatSettingsToUi(settings);
            formatTSqlExample();
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

        private void button_SaveGoogleSheetsSettings_Click(object sender, RoutedEventArgs e)
        {
            var settings = BuildGoogleSheetsSettings();
            SettingsManager.SaveGoogleSheetsSettings(settings);
            UpdateGoogleSheetsStatus(settings.refreshToken);
            SavedMessage();
        }

        private void button_SaveUpdateSettings_Click(object sender, RoutedEventArgs e)
        {
            SettingsManager.SaveEnableUpdateChecks(EnableUpdateChecks.IsChecked.GetValueOrDefault(true));
            SavedMessage();
        }

        private void button_CheckUpdates_Click(object sender, RoutedEventArgs e)
        {
            UpdateChecker.CheckNow(AxialSqlToolsPackage.PackageInstance, ignoreSettings: true);
            UpdateUpdateStatus();
        }

        private void UpdateChecker_LastUpdateResultChanged()
        {
            try
            {
                Dispatcher.BeginInvoke(new Action(UpdateUpdateStatus));
            }
            catch
            {
            }
        }

        private void UpdateUpdateStatus()
        {
            if (UpdateCheckStatus != null)
            {
                UpdateCheckStatus.Text = UpdateChecker.LastUpdateResult;
            }
        }

        private async void button_AuthorizeGoogleSheets_Click(object sender, RoutedEventArgs e)
        {
            var settings = BuildGoogleSheetsSettings();

            if (!settings.HasClientConfiguration())
            {
                MessageBox.Show("Client ID and Client Secret are required before authorizing Google Sheets.", "Google Sheets", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                string authorizationUrl = GoogleSheetsExport.BuildAuthorizationUrl(settings);
                Process.Start(new ProcessStartInfo(authorizationUrl) { UseShellExecute = true });

                string authorizationCode = Interaction.InputBox("Paste the authorization code provided by Google after granting access.", "Google Sheets Authorization");
                if (string.IsNullOrWhiteSpace(authorizationCode))
                {
                    return;
                }

                var authResult = await GoogleSheetsExport.ExchangeAuthorizationCodeAsync(settings, authorizationCode.Trim(), CancellationToken.None);

                if (!string.IsNullOrWhiteSpace(authResult.RefreshToken))
                {
                    settings.refreshToken = authResult.RefreshToken;
                }

                SettingsManager.SaveGoogleSheetsSettings(settings);
                UpdateGoogleSheetsStatus(settings.refreshToken);
                SavedMessage();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Authorization failed: {ex.Message}", "Google Sheets", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private SettingsManager.GoogleSheetsSettings BuildGoogleSheetsSettings()
        {
            return new SettingsManager.GoogleSheetsSettings
            {
                includeSourceQuery = GoogleSheetsIncludeSourceQuery.IsChecked.GetValueOrDefault(false),
                exportBoolsAsNumbers = GoogleSheetsExportBoolsAsNumbers.IsChecked.GetValueOrDefault(false),
                defaultSpreadsheetName = GoogleSheetsDefaultSpreadsheetName.Text,
                clientId = GoogleSheetsClientId.Text,
                clientSecret = GoogleSheetsClientSecret.Password,
                refreshToken = SettingsManager.GetGoogleSheetsSettings().refreshToken
            };
        }

        private void UpdateGoogleSheetsStatus(string refreshToken)
        {
            if (string.IsNullOrWhiteSpace(refreshToken))
            {
                GoogleSheetsRefreshTokenLabel.Text = "Not authorized";
                GoogleSheetsRefreshTokenLabel.Foreground = new SolidColorBrush(Colors.DarkRed);
            }
            else
            {
                GoogleSheetsRefreshTokenLabel.Text = "Authorized";
                GoogleSheetsRefreshTokenLabel.Foreground = new SolidColorBrush(Colors.DarkGreen);
            }
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
            SettingsManager.SaveQueryHistoryStorageMode(GetSelectedQueryHistoryStorageType());

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

        private string GetSelectedQueryHistoryStorageType()
        {
            if (QueryHistoryStorageType.SelectedItem is ComboBoxItem item)
            {
                return item.Tag?.ToString() ?? QueryHistoryStorageModeDatabase;
            }

            return QueryHistoryStorageModeDatabase;
        }

        private void SelectQueryHistoryStorageType(string storageType)
        {
            string mode = string.IsNullOrWhiteSpace(storageType) ? QueryHistoryStorageModeDatabase : storageType;

            foreach (var obj in QueryHistoryStorageType.Items)
            {
                if (obj is ComboBoxItem item && string.Equals(item.Tag?.ToString(), mode, StringComparison.OrdinalIgnoreCase))
                {
                    QueryHistoryStorageType.SelectedItem = item;
                    return;
                }
            }

            QueryHistoryStorageType.SelectedIndex = 0;
        }

        private void UpdateQueryHistoryStorageControls()
        {
            bool isDisabledStorage = string.Equals(GetSelectedQueryHistoryStorageType(), QueryHistoryStorageModeDisabled, StringComparison.OrdinalIgnoreCase);
            bool isDatabaseStorage = string.Equals(GetSelectedQueryHistoryStorageType(), QueryHistoryStorageModeDatabase, StringComparison.OrdinalIgnoreCase);
            Label_QueryHistoryConnectionInfoTitle.Visibility = isDatabaseStorage ? Visibility.Visible : Visibility.Collapsed;
            Label_QueryHistoryConnectionInfo.Visibility = isDatabaseStorage ? Visibility.Visible : Visibility.Collapsed;
            button_SelectDatabaseFromObjectExplorer.Visibility = isDatabaseStorage ? Visibility.Visible : Visibility.Collapsed;
            Label_QueryHistoryTargetTableName.Visibility = isDatabaseStorage ? Visibility.Visible : Visibility.Collapsed;
            QueryHistoryTableName.Visibility = isDatabaseStorage ? Visibility.Visible : Visibility.Collapsed;
            Label_QueryHistoryTargetTableHint.Visibility = isDatabaseStorage ? Visibility.Visible : Visibility.Collapsed;
            Group_QueryHistoryCreateScript.Visibility = isDatabaseStorage ? Visibility.Visible : Visibility.Collapsed;
            QueryHistoryTextFilesPanel.Visibility = (!isDatabaseStorage && !isDisabledStorage) ? Visibility.Visible : Visibility.Collapsed;
            Label_QueryHistoryTextFilesInfo.Visibility = (!isDatabaseStorage && !isDisabledStorage) ? Visibility.Visible : Visibility.Collapsed;
        }

        private void QueryHistoryStorageType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateQueryHistoryStorageControls();
        }

        private void Button_OpenQueryHistoryFolder_Click(object sender, RoutedEventArgs e)
        {
            string folderPath = SettingsManager.GetQueryHistoryTextFileFolder();
            Directory.CreateDirectory(folderPath);
            Process.Start(new ProcessStartInfo
            {
                FileName = folderPath,
                UseShellExecute = true
            });
        }

        private SettingsManager.TSqlCodeFormatSettings BuildTSqlCodeFormatSettingsFromUi()
        {
            return new SettingsManager.TSqlCodeFormatSettings
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
                leadingCommas = LeadingCommas.IsChecked.GetValueOrDefault(false),
                semicolonBeforeCte = SemicolonBeforeCte.IsChecked.GetValueOrDefault(false),
                useFreehandFormatMode = FormatModeFreehand.IsChecked.GetValueOrDefault(false),
                breakSelectElementsPerLine = BreakSelectElementsPerLine.IsChecked.GetValueOrDefault(false),
                useAssignmentAliases = UseAssignmentAliases.IsChecked.GetValueOrDefault(false),
                omitAsForTableAliases = OmitAsForTableAliases.IsChecked.GetValueOrDefault(false),
                omitAsInDeclare = OmitAsInDeclare.IsChecked.GetValueOrDefault(false),
                formatTableDefinitionsMultiline = FormatTableDefinitionsMultiline.IsChecked.GetValueOrDefault(false),
                prefixUnicodeStrings = PrefixUnicodeStrings.IsChecked.GetValueOrDefault(false),
                removeSemicolonsFromDeclare = FormatTableDefinitionsMultiline.IsChecked.GetValueOrDefault(false),
                // breakSelectFieldsAfterTopAndUnindent = BreakSelectFieldsAfterTopAndUnindent.IsChecked.GetValueOrDefault(false)
            };
        }

        private void ApplyTSqlCodeFormatSettingsToUi(SettingsManager.TSqlCodeFormatSettings settings)
        {
            PreserveComments.IsChecked = settings.preserveComments;
            RemoveNewLineAfterJoin.IsChecked = settings.removeNewLineAfterJoin;
            AddTabAfterJoinOn.IsChecked = settings.addTabAfterJoinOn;
            MoveCrossJoinToNewLine.IsChecked = settings.moveCrossJoinToNewLine;
            FormatCaseAsMultiline.IsChecked = settings.formatCaseAsMultiline;
            AddNewLineBetweenStatementsInBlocks.IsChecked = settings.addNewLineBetweenStatementsInBlocks;
            BreakSprocParametersPerLine.IsChecked = settings.breakSprocParametersPerLine;
            UppercaseBuiltInFunctions.IsChecked = settings.uppercaseBuiltInFunctions;
            UnindentBeginEndBlocks.IsChecked = settings.unindentBeginEndBlocks;
            BreakVariableDefinitionsPerLine.IsChecked = settings.breakVariableDefinitionsPerLine;
            BreakSprocDefinitionParametersPerLine.IsChecked = settings.breakSprocDefinitionParametersPerLine;
            LeadingCommas.IsChecked = settings.leadingCommas;
            SemicolonBeforeCte.IsChecked = settings.semicolonBeforeCte;
            BreakSelectElementsPerLine.IsChecked = settings.breakSelectElementsPerLine;
            UseAssignmentAliases.IsChecked = settings.useAssignmentAliases;
            OmitAsForTableAliases.IsChecked = settings.omitAsForTableAliases;
            OmitAsInDeclare.IsChecked = settings.omitAsInDeclare;
            FormatTableDefinitionsMultiline.IsChecked = settings.formatTableDefinitionsMultiline;
            PrefixUnicodeStrings.IsChecked = settings.prefixUnicodeStrings;
        }

        private SettingsManager.TSqlCodeFormatSettings LearnTSqlFormatSettingsFromSample(string formattedSample)
        {
            string sql = formattedSample ?? string.Empty;
            string normalized = NormalizeLineEndings(sql);
            string codeOnly = StripCommentsForLearning(normalized);

            return new SettingsManager.TSqlCodeFormatSettings
            {
                preserveComments = Regex.IsMatch(normalized, @"(?m)^\s*(--|/\*)"),
                removeNewLineAfterJoin = Regex.IsMatch(codeOnly, @"(?is)\bJOIN\s+(?!\r?\n)\S.+?\s+ON\b"),
                addTabAfterJoinOn = Regex.IsMatch(codeOnly, @"(?im)^\s{4,}ON\b"),
                moveCrossJoinToNewLine = Regex.IsMatch(codeOnly, @"(?im)^\s*(CROSS\s+JOIN|OUTER\s+APPLY|CROSS\s+APPLY)\b"),
                formatCaseAsMultiline = Regex.IsMatch(codeOnly, @"(?is)\bCASE\s*\r?\n\s+WHEN\b")
                    || Regex.IsMatch(codeOnly, @"(?is)\bWHEN\b.+?\r?\n\s+THEN\b"),
                addNewLineBetweenStatementsInBlocks = Regex.IsMatch(codeOnly, @";\s*(\r?\n){2,}\s*(SELECT|EXEC|DECLARE|IF|BEGIN|END)\b", RegexOptions.IgnoreCase),
                breakSprocParametersPerLine = Regex.IsMatch(codeOnly, @"(?is)\bEXEC(?:UTE)?\s+[\[\]\w.]+\s*\r?\n\s*,?\s*@\w+.+\r?\n\s*,?\s*@\w+"),
                uppercaseBuiltInFunctions = Regex.IsMatch(codeOnly, @"\b(GETDATE|DATEADD|COUNT|SUM)\s*\(")
                    && !Regex.IsMatch(codeOnly, @"\b(getdate|dateadd|count|sum)\s*\("),
                unindentBeginEndBlocks = Regex.IsMatch(codeOnly, @"(?m)^BEGIN\s*$")
                    && Regex.IsMatch(codeOnly, @"(?m)^END\s*$"),
                breakVariableDefinitionsPerLine = Regex.IsMatch(codeOnly, @"(?is)\bDECLARE\s+@\w+.+\r?\n\s*,?\s*@\w+"),
                breakSprocDefinitionParametersPerLine = Regex.IsMatch(codeOnly, @"(?is)\bPROCEDURE\s+[\[\]\w.]+\s*\r?\n\s*,?\s*@\w+.+\r?\n\s*,?\s*@\w+"),
                leadingCommas = Regex.IsMatch(codeOnly, @"(?m)^\s*,\s*\S"),
                semicolonBeforeCte = Regex.IsMatch(codeOnly, @"(?im)^\s*;\s*WITH\b"),
                breakSelectElementsPerLine = Regex.IsMatch(codeOnly, @"(?im)^\s*SELECT(?:\s+TOP\s+\d+)?\s*$")
                    || Regex.IsMatch(codeOnly, @"(?m)^\s*,\s*[\w@]"),
                useAssignmentAliases = Regex.IsMatch(codeOnly, @"(?m)^\s*,?\s*\w+\s*=\s*"),
                omitAsForTableAliases = Regex.IsMatch(codeOnly, @"(?i)\b(FROM|JOIN|APPLY)\s+@\w+\s+\w+\b")
                    && !Regex.IsMatch(codeOnly, @"(?i)\b(FROM|JOIN|APPLY)\s+@\w+\s+AS\s+\w+\b"),
                omitAsInDeclare = Regex.IsMatch(codeOnly, @"(?im)^\s*DECLARE\s+@\w+\s+(?!AS\b)\w+"),
                formatTableDefinitionsMultiline = Regex.IsMatch(codeOnly, @"(?is)\bDECLARE\s+@\w+\s+TABLE\s*\(\s*\r?\n"),
                prefixUnicodeStrings = Regex.IsMatch(codeOnly, @"(?<![A-Za-z])N'"),
                removeSemicolonsFromDeclare = Regex.IsMatch(codeOnly, @"(?im)^\s*DECLARE\s+@\w+.*[^;]\s*$")
            };
        }

        private static string StripCommentsForLearning(string sql)
        {
            string withoutBlockComments = Regex.Replace(sql ?? string.Empty, @"/\*.*?\*/", string.Empty, RegexOptions.Singleline);
            return Regex.Replace(withoutBlockComments, @"(?m)--.*$", string.Empty);
        }

        private bool ValidateFreehandFormatSample()
        {
            try
            {
                var learnedSettings = LearnTSqlFormatSettingsFromSample(FreehandFormatEditor.Text);
                string learnedSourceFormat = TSqlFormatter.FormatCode(FreehandSourceEditor.Text, learnedSettings);
                string sourceCanonical = CanonicalizeSqlForEquivalence(learnedSourceFormat);
                string formattedCanonical = CanonicalizeSqlForEquivalence(FreehandFormatEditor.Text);

                if (string.Equals(sourceCanonical, formattedCanonical, StringComparison.Ordinal))
                {
                    return true;
                }

                MessageBox.Show(
                    "The formatted sample must contain the same SQL as the source sample. You can change whitespace, casing, comments, comma placement, and semicolon placement, but not add, remove, or rewrite SQL.",
                    "Format sample changed SQL",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"The formatted sample could not be parsed as valid T-SQL:{Environment.NewLine}{ex.Message}",
                    "Format sample parse error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }

            return false;
        }

        private static string CanonicalizeSqlForEquivalence(string sql)
        {
            TSql170Parser sqlParser = new TSql170Parser(false);
            IList<ParseError> parseErrors;
            TSqlFragment fragment = sqlParser.Parse(new StringReader(sql ?? string.Empty), out parseErrors);

            if (parseErrors.Count > 0)
            {
                throw new Exception(BuildParseErrorMessage(parseErrors));
            }

            StringBuilder result = new StringBuilder();
            foreach (var token in fragment.ScriptTokenStream)
            {
                if (!IsSignificantSqlToken(token))
                {
                    continue;
                }

                if (result.Length > 0)
                {
                    result.Append('\n');
                }

                result.Append(NormalizeSqlTokenType(token));
                result.Append(':');
                result.Append(NormalizeSqlTokenText(token));
            }

            return result.ToString();
        }

        private static bool IsSignificantSqlToken(TSqlParserToken token)
        {
            if (token == null || string.IsNullOrEmpty(token.Text))
            {
                return false;
            }

            return token.TokenType != TSqlTokenType.WhiteSpace
                && token.TokenType != TSqlTokenType.SingleLineComment
                && token.TokenType != TSqlTokenType.MultilineComment
                && token.TokenType != TSqlTokenType.Semicolon
                && token.Text != ";";
        }

        private static string NormalizeSqlTokenType(TSqlParserToken token)
        {
            if (token.TokenType == TSqlTokenType.AsciiStringLiteral
                || token.TokenType == TSqlTokenType.UnicodeStringLiteral)
            {
                return "StringLiteral";
            }

            return token.TokenType.ToString();
        }

        private static string NormalizeSqlTokenText(TSqlParserToken token)
        {
            if (token.TokenType == TSqlTokenType.UnicodeStringLiteral
                && token.Text.Length > 1
                && char.ToUpperInvariant(token.Text[0]) == 'N'
                && token.Text[1] == '\'')
            {
                return token.Text.Substring(1);
            }

            if (token.TokenType == TSqlTokenType.AsciiStringLiteral
                || token.TokenType == TSqlTokenType.UnicodeStringLiteral
                || token.TokenType == TSqlTokenType.Integer
                || token.TokenType == TSqlTokenType.Real
                || token.TokenType == TSqlTokenType.HexLiteral)
            {
                return token.Text;
            }

            return token.Text.ToUpperInvariant();
        }

        private static string BuildParseErrorMessage(IList<ParseError> parseErrors)
        {
            StringBuilder errorBuilder = new StringBuilder();
            foreach (var parseError in parseErrors)
            {
                errorBuilder.AppendLine($"Line {parseError.Line}, column {parseError.Column}: {parseError.Message}");
            }

            return errorBuilder.ToString().Trim();
        }

        private static string NormalizeLineEndings(string text)
        {
            return (text ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "\r\n");
        }

        private void UpdateFormatModeUi()
        {
            if (FormatOptionsPanel == null || FreehandFormatPanel == null || FormatPreviewPanel == null)
            {
                return;
            }

            bool isFreehand = FormatModeFreehand.IsChecked == true;
            FormatOptionsPanel.Visibility = isFreehand ? Visibility.Collapsed : Visibility.Visible;
            FormatPreviewPanel.Visibility = isFreehand ? Visibility.Collapsed : Visibility.Visible;
            FreehandFormatPanel.Visibility = isFreehand ? Visibility.Visible : Visibility.Collapsed;
        }

        private void formatTSqlExample()
        {
            if (SourceQueryPreview == null || FormattedQueryPreview == null)
            {
                return;
            }

            var settings = BuildTSqlCodeFormatSettingsFromUi();

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

        private void FormatMode_Checked(object sender, RoutedEventArgs e)
        {
            UpdateFormatModeUi();
        }

        private void Button_FormatFreehandSample_Click(object sender, RoutedEventArgs e)
        {
            FreehandFormatEditor.Text = TSqlFormatter.FormatCode(FreehandSourceEditor.Text, SettingsManager.GetTSqlCodeFormatSettings());
        }

        private void Button_OpenFreehandSampleInNewTab_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ServiceCache.ScriptFactory.CreateNewBlankScript(ScriptType.Sql);

                EnvDTE.DTE dte = Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(EnvDTE.DTE)) as EnvDTE.DTE;
                if (dte?.ActiveDocument == null)
                {
                    return;
                }

                EnvDTE.TextSelection selection = dte.ActiveDocument.Selection as EnvDTE.TextSelection;
                if (selection != null)
                {
                    selection.Insert(FreehandFormatEditor.Text);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to open the sample in a new SQL tab: {ex.Message}", "Open SQL tab");
            }
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

        private void LoadConnectionColorRules()
        {
            _connectionColorRules = new ObservableCollection<SettingsManager.ConnectionColorRule>(SettingsManager.GetConnectionColorRules());
            ConnectionColorRulesListView.ItemsSource = _connectionColorRules;
        }

        private void EnsureConnectionColorRulesLoaded()
        {
            if (_connectionColorRules == null)
            {
                _connectionColorRules = new ObservableCollection<SettingsManager.ConnectionColorRule>();
                ConnectionColorRulesListView.ItemsSource = _connectionColorRules;
            }
        }

        private string PickColor(string currentHex)
        {
            var dialog = new System.Windows.Forms.ColorDialog();
            try
            {
                dialog.Color = System.Drawing.ColorTranslator.FromHtml(currentHex);
            }
            catch { }
            dialog.FullOpen = true;

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                return System.Drawing.ColorTranslator.ToHtml(dialog.Color);
            }
            return null;
        }

        private void NewRuleColorPreview_Click(object sender, RoutedEventArgs e)
        {
            var currentBrush = NewRuleColorPreview.Background as SolidColorBrush;
            string currentHex = currentBrush != null
                ? string.Format("#{0:X2}{1:X2}{2:X2}", currentBrush.Color.R, currentBrush.Color.G, currentBrush.Color.B)
                : "#FF4444";

            string picked = PickColor(currentHex);
            if (picked != null)
            {
                try
                {
                    var color = System.Drawing.ColorTranslator.FromHtml(picked);
                    NewRuleColorPreview.Background = new SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(color.R, color.G, color.B));
                }
                catch { }
            }
        }

        private void NewRuleColorPreview_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            NewRuleColorPreview_Click(sender, (RoutedEventArgs)e);
        }

        private void RuleColorPreview_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.DataContext is SettingsManager.ConnectionColorRule rule)
            {
                string picked = PickColor(rule.StatusBarColor);
                if (picked != null)
                {
                    rule.StatusBarColor = picked;
                    ConnectionColorRulesListView.Items.Refresh();
                }
            }
        }

        private void ButtonAddColorRule_Click(object sender, RoutedEventArgs e)
        {
            EnsureConnectionColorRulesLoaded();

            string serverPattern = NewRuleServerPattern.Text?.Trim();
            string databasePattern = NewRuleDatabasePattern.Text?.Trim();

            if (string.IsNullOrEmpty(serverPattern) && string.IsNullOrEmpty(databasePattern))
            {
                MessageBox.Show("Fill in at least the server name or the database name.", "Connection Colors");
                return;
            }

            var brush = NewRuleColorPreview.Background as SolidColorBrush;
            string hex = "#FF4444";
            if (brush != null)
            {
                hex = string.Format("#{0:X2}{1:X2}{2:X2}", brush.Color.R, brush.Color.G, brush.Color.B);
            }

            _connectionColorRules.Add(new SettingsManager.ConnectionColorRule
            {
                ServerNamePattern = serverPattern ?? string.Empty,
                DatabaseNamePattern = databasePattern ?? string.Empty,
                StatusBarColor = hex,
                IsEnabled = true
            });

            NewRuleServerPattern.Text = "";
            NewRuleDatabasePattern.Text = "";
        }

        private void ButtonEditColorRule_Click(object sender, RoutedEventArgs e)
        {
            if (ConnectionColorRulesListView.SelectedItem is SettingsManager.ConnectionColorRule selectedRule)
            {
                NewRuleServerPattern.Text = selectedRule.ServerNamePattern;
                NewRuleDatabasePattern.Text = selectedRule.DatabaseNamePattern;

                try
                {
                    var color = System.Drawing.ColorTranslator.FromHtml(selectedRule.StatusBarColor);
                    NewRuleColorPreview.Background = new SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(color.R, color.G, color.B));
                }
                catch { }

                _connectionColorRules.Remove(selectedRule);
            }
        }

        private void ButtonRemoveColorRule_Click(object sender, RoutedEventArgs e)
        {
            if (ConnectionColorRulesListView.SelectedItem is SettingsManager.ConnectionColorRule selectedRule)
            {
                _connectionColorRules.Remove(selectedRule);
            }
        }

        private void Button_SaveConnectionColorRules_Click(object sender, RoutedEventArgs e)
        {
            var rules = new System.Collections.Generic.List<SettingsManager.ConnectionColorRule>(_connectionColorRules);
            SettingsManager.SaveConnectionColorRules(rules);
            SavedMessage();
        }

    }
}
