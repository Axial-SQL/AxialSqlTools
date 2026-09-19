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
        private bool _settingsLoaded;
        private bool _authorizingGoogleSheets;
        private int _authorizationVersion;
        private string _savedGitHubToken = string.Empty;
        private SettingsManager.GoogleSheetsSettings _googleSheetsDraft;
        private readonly ToolWindowThemeController _themeController;
        private bool updateResultSubscribed;
        private ObservableCollection<SettingsManager.ConnectionColorRule> _connectionColorRules;

        private string tsqlFormatExample = @"while (1=0) 
begin -- inline comment
select distinct top 10
    c.CustomerID, getDate(),
    CASE WHEN o.TotalAmount > 1000 THEN 'High' ELSE 'Low' END AS OrderSize
FROM Customers c
JOIN Orders o ON c.CustomerID = o.CustomerID CROSS JOIN Regions r
WHERE c.IsActive = 1; /* multi-line 
comment */
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

            _connectionColorRules = new ObservableCollection<SettingsManager.ConnectionColorRule>();
            ConnectionColorRulesListView.ItemsSource = _connectionColorRules;

            _themeController = new ToolWindowThemeController(this, ApplyThemeBrushResources);

            this.Loaded += UserControl_Loaded;
            this.Unloaded += UserControl_Unloaded;

            SourceQueryPreview.Text = tsqlFormatExample;

            formatTSqlExample();

        }

        private void UserControl_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            SubscribeToUpdateResultChanges();
            if (!_settingsLoaded)
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

            bool isAuthorized = GoogleSheetsRefreshTokenLabel.Text.StartsWith("Authorized", StringComparison.OrdinalIgnoreCase);
            GoogleSheetsRefreshTokenLabel.Foreground = GetThemedStatusBrush(isAuthorized);
        }

        private void LoadSavedSettings()
        {
            _authorizationVersion++;
            _settingsLoaded = false;
            SaveAllButton.IsEnabled = false;
            try
            {

                var generalSettings = SettingsManager.GetGeneralSettings();
                UseTransactionWarning.IsChecked = generalSettings.useTransactionWarning;
                UseAlwaysEncryptedWarning.IsChecked = generalSettings.useAlwaysEncryptedWarning;
                UsePreciseExecutionTime.IsChecked = generalSettings.usePreciseExecutionTime;
                AlignNumericValuesToRight.IsChecked = generalSettings.alignNumericValuesToRight;

                ScriptFolder.Text = SettingsManager.GetTemplatesFolder();
                UpdateTemplatesFolderStatus();

                var snippetSettings = SettingsManager.GetSnippetSettings();
                UseSnippets.IsChecked = snippetSettings.useSnippets;
                SnippetFolder.Text = snippetSettings.snippetFolder;
                SnippetReplaceMode.SelectedValue = snippetSettings.replaceKey.ToString();
                var asteriskSettings = SettingsManager.GetAsteriskExpansionSettings();
                UseAsteriskExpansion.IsChecked = asteriskSettings.useAsteriskExpansion;
                AsteriskExpansionTriggerMode.SelectedValue = asteriskSettings.triggerKey.ToString();

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
                DisregardSsmsFormatterSettings.IsChecked = tsqlCodeFormatSettings.disregardSsmsFormatterSettings;
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
                BreakSelectFieldsAfterTopAndUnindent.IsChecked = tsqlCodeFormatSettings.breakSelectFieldsAfterTopAndUnindent;

                // Excel export settings
                var excelSettings = SettingsManager.GetExcelExportSettings();
                ExcelExportIncludeSourceQuery.IsChecked = excelSettings.includeSourceQuery;
                ExcelExportAddAutoFilter.IsChecked = excelSettings.addAutofilter;
                ExcelExportBoolsAsNumbers.IsChecked = excelSettings.exportBoolsAsNumbers;
                ExcelExportDefaultDirectory.Text = excelSettings.defaultDirectory;
                ExcelExportDefaultFilename.Text = excelSettings.defaultFileName;

                var googleSettings = SettingsManager.GetGoogleSheetsSettings();
                _googleSheetsDraft = googleSettings;
                GoogleSheetsIncludeSourceQuery.IsChecked = googleSettings.includeSourceQuery;
                GoogleSheetsExportBoolsAsNumbers.IsChecked = googleSettings.exportBoolsAsNumbers;
                GoogleSheetsDefaultSpreadsheetName.Text = googleSettings.defaultSpreadsheetName;
                GoogleSheetsClientId.Text = googleSettings.clientId;
                GoogleSheetsClientSecret.Password = googleSettings.clientSecret;
                UpdateGoogleSheetsStatus(googleSettings.refreshToken);

                WarnWhenRunningFatalAction.IsChecked = SettingsManager.GetWarnWhenRunningFatalAction();
                EnableUpdateChecks.IsChecked = SettingsManager.GetEnableUpdateChecks();
                UpdateUpdateStatus();

                LoadConnectionColorRules();
                NewRuleServerPattern.Clear();
                NewRuleDatabasePattern.Clear();
                NewRuleColorPreview.Background = new SolidColorBrush(Color.FromRgb(0xFF, 0x44, 0x44));
                SourceQueryPreview.Text = tsqlFormatExample;
                formatTSqlExample();
                _settingsLoaded = true;

            }
            catch (Exception ex)
            {
                _logger.Error(ex, "An exception occurred while loading settings");

                string msg = $"Error message: {ex.Message} \nInnerException: {ex.InnerException}";
                MessageBox.Show(msg, "Error");
            }

            GitHubToken.Password = string.Empty;
            try
            {
                GitHubToken.Password = WindowsCredentialHelper.LoadToken("AxialSqlTools_GitHubToken");
            }
            catch 
            {
                // No credential has been saved, or Credential Manager is unavailable.
            }
            _savedGitHubToken = GitHubToken.Password;
            SaveAllButton.IsEnabled = _settingsLoaded && !_authorizingGoogleSheets;

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

        private SettingsManager.SnippetReplaceKey GetSelectedAsteriskExpansionTriggerKey()
        {
            var selectedValue = AsteriskExpansionTriggerMode.SelectedValue as string;
            if (Enum.TryParse(selectedValue, out SettingsManager.SnippetReplaceKey key))
            {
                return key;
            }

            return SettingsManager.SnippetReplaceKey.Tab;
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

        private void SaveAllSettings_Click(object sender, RoutedEventArgs e)
        {
            if (!_settingsLoaded || _authorizingGoogleSheets)
                return;

            if (!int.TryParse(SMTP_Port.Text, out int smtpPort) || smtpPort < 1 || smtpPort > 65535)
            {
                SettingsTabs.SelectedItem = SmtpSettingsTab;
                MessageBox.Show("Enter an SMTP port between 1 and 65535.", "Invalid settings",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                SMTP_Port.Focus();
                SMTP_Port.SelectAll();
                return;
            }

            var snippets = SettingsManager.GetSnippetSettings();
            snippets.useSnippets = UseSnippets.IsChecked.GetValueOrDefault();
            snippets.snippetFolder = SnippetFolder.Text;
            snippets.replaceKey = GetSelectedSnippetReplaceKey();

            var settings = new SettingsManager.WindowSettings
            {
                General = new SettingsManager.GeneralSettings
                {
                    useTransactionWarning = UseTransactionWarning.IsChecked.GetValueOrDefault(),
                    useAlwaysEncryptedWarning = UseAlwaysEncryptedWarning.IsChecked.GetValueOrDefault(),
                    usePreciseExecutionTime = UsePreciseExecutionTime.IsChecked.GetValueOrDefault(),
                    alignNumericValuesToRight = AlignNumericValuesToRight.IsChecked.GetValueOrDefault()
                },
                TemplatesFolder = ScriptFolder.Text,
                Snippets = snippets,
                AsteriskExpansion = new SettingsManager.AsteriskExpansionSettings
                {
                    useAsteriskExpansion = UseAsteriskExpansion.IsChecked.GetValueOrDefault(),
                    triggerKey = GetSelectedAsteriskExpansionTriggerKey()
                },
                QueryHistoryConnectionString = _queryHistoryConnectionString,
                QueryHistoryTableName = QueryHistoryTableName.Text,
                QueryHistoryStorageMode = GetSelectedQueryHistoryStorageType(),
                MyEmail = MyEmailAddress.Text,
                Smtp = new SettingsManager.SmtpSettings
                {
                    ServerName = SMTP_Server.Text,
                    Port = smtpPort,
                    Username = SMTP_UserName.Text,
                    Password = SMTP_Password.Password,
                    EnableSsl = SMTP_EnableSSL.IsChecked.GetValueOrDefault()
                },
                CodeFormat = BuildCodeFormatSettings(),
                ExcelExport = new SettingsManager.ExcelExportSettings
                {
                    includeSourceQuery = ExcelExportIncludeSourceQuery.IsChecked.GetValueOrDefault(),
                    addAutofilter = ExcelExportAddAutoFilter.IsChecked.GetValueOrDefault(),
                    exportBoolsAsNumbers = ExcelExportBoolsAsNumbers.IsChecked.GetValueOrDefault(),
                    defaultDirectory = ExcelExportDefaultDirectory.Text,
                    defaultFileName = ExcelExportDefaultFilename.Text
                },
                GoogleSheets = BuildGoogleSheetsSettings(),
                ConnectionColorRules = new System.Collections.Generic.List<SettingsManager.ConnectionColorRule>(_connectionColorRules),
                WarnWhenRunningFatalAction = WarnWhenRunningFatalAction.IsChecked == true,
                EnableUpdateChecks = EnableUpdateChecks.IsChecked.GetValueOrDefault(true)
            };

            if (!SaveSettings(() => SettingsManager.SaveWindowSettings(settings)))
                return;

            _googleSheetsDraft = settings.GoogleSheets;
            UpdateGoogleSheetsStatus(_googleSheetsDraft.refreshToken);
            UpdateTemplatesFolderStatus();
            RefreshQueryHistoryCreateScript();
            GridAccess.ColorAllDocumentTabs();
            GridAccess.ScheduleReapplyAllTabColors();

            // The hidden GitHub tab uses Credential Manager, not settings.json.
            // Leave its credential alone unless it was explicitly edited.
            if (!string.Equals(GitHubToken.Password, _savedGitHubToken, StringComparison.Ordinal))
            {
                try
                {
                    WindowsCredentialHelper.SaveToken("AxialSqlTools_GitHubToken", "AxialSqlTools_GitHubToken", GitHubToken.Password);
                    _savedGitHubToken = GitHubToken.Password;
                }
                catch (Exception ex)
                {
                    SettingsFileStore.ReportSaveFailure(ex);
                    MessageBox.Show("Settings were saved, but the GitHub token could not be saved to Windows Credential Manager. Please try again.",
                        "GitHub token not saved", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }

            SavedMessage();
        }

        private void DiscardChanges_Click(object sender, RoutedEventArgs e)
        {
            LoadSavedSettings();
        }

        private bool SaveSettings(Func<bool> save)
        {
            try
            {
                if (save())
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                SettingsFileStore.ReportSaveFailure(ex);
            }

            MessageBox.Show(SettingsManager.LastSaveError ?? "The settings could not be saved. Please try again.",
                "Settings not saved", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }

        private void UpdateTemplatesFolderStatus()
        {
            bool available = Directory.Exists(SettingsManager.GetTemplatesFolder());
            TemplatesFolderStatus.Text = available ? string.Empty
                : "The templates folder is currently unavailable or has not been created. The configured path is retained.";
            TemplatesFolderStatus.Visibility = available ? Visibility.Collapsed : Visibility.Visible;
        }

        private void SavedMessage()
        {
            MessageBox.Show(
                "All settings have been saved.",
                "Settings saved");
        }

        private void buttonWikiPage_Click(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
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
            if (_authorizingGoogleSheets) return;
            var settings = BuildGoogleSheetsSettings();
            int authorizationVersion = _authorizationVersion;

            if (!settings.HasClientConfiguration())
            {
                MessageBox.Show("Client ID and Client Secret are required before authorizing Google Sheets.", "Google Sheets", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _authorizingGoogleSheets = true;
            button_AuthorizeGoogleSheets.IsEnabled = false;
            SaveAllButton.IsEnabled = false;
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

                if (authorizationVersion != _authorizationVersion
                    || !string.Equals(settings.clientId, GoogleSheetsClientId.Text, StringComparison.Ordinal)
                    || !string.Equals(settings.clientSecret, GoogleSheetsClientSecret.Password, StringComparison.Ordinal))
                    return;

                if (!string.IsNullOrWhiteSpace(authResult.RefreshToken))
                {
                    settings.refreshToken = authResult.RefreshToken;
                }

                _googleSheetsDraft = settings;
                UpdateGoogleSheetsStatus(settings.refreshToken);
                if (!string.IsNullOrWhiteSpace(settings.refreshToken))
                    GoogleSheetsRefreshTokenLabel.Text = "Authorized (click Save to keep changes)";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Authorization failed: {ex.Message}", "Google Sheets", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _authorizingGoogleSheets = false;
                button_AuthorizeGoogleSheets.IsEnabled = true;
                SaveAllButton.IsEnabled = _settingsLoaded;
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
                refreshToken = _googleSheetsDraft != null
                    && string.Equals(_googleSheetsDraft.clientId, GoogleSheetsClientId.Text, StringComparison.Ordinal)
                    && string.Equals(_googleSheetsDraft.clientSecret, GoogleSheetsClientSecret.Password, StringComparison.Ordinal)
                        ? _googleSheetsDraft.refreshToken : string.Empty
            };
        }

        private void GoogleSheetsCredentialsChanged(object sender, RoutedEventArgs e)
        {
            if (_settingsLoaded)
                UpdateGoogleSheetsStatus(BuildGoogleSheetsSettings().refreshToken);
        }

        private void UpdateGoogleSheetsStatus(string refreshToken)
        {
            if (string.IsNullOrWhiteSpace(refreshToken))
            {
                GoogleSheetsRefreshTokenLabel.Text = "Not authorized";
                GoogleSheetsRefreshTokenLabel.Foreground = GetThemedStatusBrush(false);
            }
            else
            {
                GoogleSheetsRefreshTokenLabel.Text = "Authorized";
                GoogleSheetsRefreshTokenLabel.Foreground = GetThemedStatusBrush(true);
            }
        }

        private void Hyperlink_RequestNavigateFormatQueryWiki(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
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

        private SettingsManager.TSqlCodeFormatSettings BuildCodeFormatSettings()
        {
            return new SettingsManager.TSqlCodeFormatSettings
            {
                disregardSsmsFormatterSettings = DisregardSsmsFormatterSettings.IsChecked.GetValueOrDefault(false),
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
                breakSelectFieldsAfterTopAndUnindent = BreakSelectFieldsAfterTopAndUnindent.IsChecked.GetValueOrDefault(false)
            };

        }

        private int _formatterPreviewRequest;
        private bool _formatterPreviewDisregardSsmsSettings;
        private System.Threading.Tasks.Task<SsmsFormatterContext> _formatterPreviewContextTask;

        private async void formatTSqlExample()
        {
            // Checked events can fire while InitializeComponent is still building the controls.
            if (SourceQueryPreview == null || FormattedQueryPreview == null
                || BreakSelectFieldsAfterTopAndUnindent == null || DisregardSsmsFormatterSettings == null) return;
            int request = ++_formatterPreviewRequest;
            try
            {
                string source = SourceQueryPreview.Text;
                var settings = BuildCodeFormatSettings();
                // Coalesce checkbox events raised together when settings are loaded/discarded.
                if (_formatterPreviewContextTask == null || _formatterPreviewContextTask.IsCompleted
                    || _formatterPreviewDisregardSsmsSettings != settings.disregardSsmsFormatterSettings)
                {
                    _formatterPreviewDisregardSsmsSettings = settings.disregardSsmsFormatterSettings;
                    _formatterPreviewContextTask = SsmsFormatterHost.CreateAsync(null,
                        System.Threading.CancellationToken.None, settings.disregardSsmsFormatterSettings);
                }
                var formatter = await _formatterPreviewContextTask;
                if (request != _formatterPreviewRequest) return;
                FormattedQueryPreview.Text = TSqlFormatter.FormatCode(source, settings, formatter.Parser, formatter.Generator);
            }
            catch (Exception ex)
            {
                if (request == _formatterPreviewRequest)
                    FormattedQueryPreview.Text = "Preview unavailable: " + ex.Message;
            }
        }

        private void formatSetting_Checked(object sender, RoutedEventArgs e)
        {
            formatTSqlExample();
        }

        private void formatSetting_Unchecked(object sender, RoutedEventArgs e)
        {
            formatTSqlExample();
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

    }
}
