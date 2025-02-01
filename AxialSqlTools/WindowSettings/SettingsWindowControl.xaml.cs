namespace AxialSqlTools
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.IO.Compression;
    using System.Net.Http;
    using System.Windows;
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for SettingsWindowControl.
    /// </summary>
    public partial class SettingsWindowControl : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SettingsWindowControl"/> class.
        /// </summary>
        public SettingsWindowControl()
        {
            this.InitializeComponent();

            LoadSavedSettings();

            this.Loaded += UserControl_Loaded;

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

                QueryHistoryConnectionString.Text = SettingsManager.GetQueryHistoryConnectionString();
                QueryHistoryTableName.Text = SettingsManager.GetQueryHistoryTableName();

                MyEmailAddress.Text = SettingsManager.GetMyEmail();

                SettingsManager.SmtpSettings smtpSettings = SettingsManager.GetSmtpSettings();

                SMTP_Server.Text = smtpSettings.ServerName;
                SMTP_Port.Text = smtpSettings.Port.ToString();
                SMTP_UserName.Text = smtpSettings.Username;
                SMTP_Password.Password = smtpSettings.Password;
                SMTP_EnableSSL.IsChecked = smtpSettings.EnableSsl;

                ApplyCodeFormat.IsChecked = SettingsManager.GetApplyAdditionalCodeFormatting();

                OpenAiApiKey.Password = SettingsManager.GetOpenAiApiKey();

            }
            catch (Exception ex)
            {
                string msg = $"Erorr message: {ex.Message} \nInnerException: {ex.InnerException}";
                MessageBox.Show(msg, "Error");
            }

        }

        ///// <summary>
        ///// Handles click on the button by displaying a message box.
        ///// </summary>
        ///// <param name="sender">The event sender.</param>
        ///// <param name="e">The event args.</param>
        //[SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Justification = "Sample code")]
        //[SuppressMessage("StyleCop.CSharp.NamingRules", "SA1300:ElementMustBeginWithUpperCaseLetter", Justification = "Default event handler naming pattern")]
        //private void button1_Click(object sender, RoutedEventArgs e)
        //{
        //    MessageBox.Show(
        //        string.Format(System.Globalization.CultureInfo.CurrentUICulture, "Invoked '{0}'", this.ToString()),
        //        "SettingsWindow");
        //}

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
            SettingsManager.SaveApplyAdditionalCodeFormatting(ApplyCodeFormat.IsChecked.GetValueOrDefault());

            SavedMessage();
        }

        private void Button_SaveOpenAi_Click(object sender, RoutedEventArgs e)
        {
            SettingsManager.SaveOpenAiApiKey(OpenAiApiKey.Password);

            SavedMessage();
        }

        private void Button_SaveQueryHistory_Click(object sender, RoutedEventArgs e)
        {
            SettingsManager.SaveQueryHistoryConnectionString(QueryHistoryConnectionString.Text);
            SettingsManager.SaveQueryHistoryTableName(QueryHistoryTableName.Text); 

            SavedMessage();
        }

        private void Button_SelectDatabaseFromObjectExplorer_Click(object sender, RoutedEventArgs e)
        {

            var ci = ScriptFactoryAccess.GetCurrentConnectionInfo();

            QueryHistoryConnectionString.Text = ci.FullConnectionString;

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
            QueryHistoryConnectionString.Text = "";
        }
    }
}