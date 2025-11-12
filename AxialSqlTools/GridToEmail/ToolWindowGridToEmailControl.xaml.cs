namespace AxialSqlTools
{
    using Microsoft.VisualStudio.Shell;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.SqlClient;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.Net;
    using System.Net.Mail;
    using System.Text;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Documents;
    using System.Windows.Input;
    using System.Windows.Media;
    using static AxialSqlTools.AxialSqlToolsPackage;


    /// <summary>
    /// Interaction logic for ToolWindowGridToEmailControl.
    /// </summary>
    public partial class ToolWindowGridToEmailControl : UserControl
    {

        private ScriptFactoryAccess.ConnectionInfo connectionInfo;
        private string exportedFilename;
        private List<DataTable> dataTables;
        private bool gridToHtmlSupported = false;

        private static string ConvertDataTableToHTML(DataTable dataTable)
        {
            StringBuilder html = new StringBuilder();

            html.Append("<table border='1' style='border-collapse: collapse;'>");

            html.Append("<thead style='background-color: grey;'><tr>");
            foreach (DataColumn column in dataTable.Columns)
            {
                string columnName = column.ColumnName;
                if (column.ExtendedProperties.ContainsKey("columnName"))
                    columnName = (string)column.ExtendedProperties["columnName"];

                html.AppendFormat("<th style='text-align: left; font-weight: bold; padding: 8px;'>{0}</th>", columnName);
            }
            html.Append("</tr></thead>");

            // Add the data rows.
            html.Append("<tbody>");
            foreach (DataRow row in dataTable.Rows)
            {
                html.Append("<tr>");
                foreach (DataColumn column in dataTable.Columns)
                {
                    // Apply right alignment for numeric values.
                    string align = column.DataType == typeof(string) ? "left" : "right";
                    html.AppendFormat("<td style='text-align: {0}; padding: 8px;'>{1}</td>", align, row[column]);
                }
                html.Append("</tr>");
            }
            html.Append("</tbody>");

            // End with the HTML table end tag.
            html.Append("</table>");

            return html.ToString();
        }    

        private string GetEmailBodyHtml()
        {
            TextRange textRange = new TextRange(EmailBody.Document.ContentStart, EmailBody.Document.ContentEnd);

            //TODO - find how to convert it to proper HTML
            var textWithNewLines = textRange.Text;

            var htmlText = textWithNewLines.Replace("\r\n", "<br>");
            htmlText = htmlText.Replace("\n", "<br>");

            if (gridToHtmlSupported && htmlText.Contains("{GRID}"))
            {
                StringBuilder htmlTables = new StringBuilder();

                foreach (DataTable dataTable in dataTables)
                {
                    htmlTables.AppendLine(ConvertDataTableToHTML(dataTable));
                    htmlTables.AppendLine("<hr/>");
                }

                htmlText = htmlText.Replace("{GRID}", htmlTables.ToString());
            }
            else
                htmlText = htmlText.Replace("{GRID}", "");

            return htmlText;
        }

        private class MailConfigItem
        {           

            public string EmailAddress { get; set; }
            public string MailProfileName { get; set; }
            public bool isSMTP { get; set; }
            public bool isDatabaseMail { get; set; }

            public MailConfigItem(string emailAddress)
            {
                isSMTP = true;
                EmailAddress = emailAddress;
            }

            public MailConfigItem(string databaseMailProfileName, string emailAddress)
            {
                isDatabaseMail = true;
                MailProfileName = databaseMailProfileName;
                EmailAddress = emailAddress;
            }

            // Override the ToString method
            public override string ToString()
            {

                string Name = "";

                if (isSMTP)
                {
                    Name = "SMTP | " + EmailAddress;
                }
                else if (isDatabaseMail)
                {
                    Name = $"Database Mail Profile '{MailProfileName}' / {EmailAddress}";
                }

                return Name; 
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ToolWindowGridToEmailControl"/> class.
        /// </summary>
        public ToolWindowGridToEmailControl()
        {
            this.InitializeComponent();
            
            // Apply theme when control loads
            this.Loaded += (s, e) =>
            {
                ApplyThemeColors();
            };
            
            // Re-apply theme when control becomes visible (e.g., switching windows or changing SSMS theme)
            this.IsVisibleChanged += (s, e) =>
            {
                if (this.IsVisible)
                {
                    ApplyThemeColors();
                }
            };
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
                    ClearThemeFromChildren(this);
                    return;
                }

                // Dark mode - apply dark theme colors
                var bgBrush = ThemeManager.GetBackgroundBrush();
                var fgBrush = ThemeManager.GetForegroundBrush();

                this.Background = bgBrush;
                this.Foreground = fgBrush;

                // Recursively apply to all children
                ApplyThemeToChildren(this, bgBrush, fgBrush);

                // Apply custom CheckBox style for dark mode
                ApplyCheckBoxStyle();
            }
            catch (Exception ex)
            {
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
                    checkBox.ClearValue(CheckBox.StyleProperty);
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
                else if (child is Image image)
                {
                    // Clear any effects applied for dark mode
                    image.Effect = null;
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
                    // Skip CheckBox - handled by custom style
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
                else if (child is Button button)
                {
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
                    grid.Background = bgBrush;
                }
                else if (child is Border border)
                {
                    if (border.Background != null)
                    {
                        border.Background = bgBrush;
                    }
                }
                else if (child is Image image)
                {
                    // Make icons visible in dark mode by adding a light background effect
                    if (ThemeManager.IsDarkTheme())
                    {
                        // Add a drop shadow or background effect to make dark icons visible
                        // Use DropShadowEffect with white color to create a glow
                        var effect = new System.Windows.Media.Effects.DropShadowEffect
                        {
                            Color = Colors.White,
                            Direction = 0,
                            ShadowDepth = 0,
                            BlurRadius = 3,
                            Opacity = 0.9
                        };
                        image.Effect = effect;
                    }
                    else
                    {
                        // Light mode: no effect needed
                        image.Effect = null;
                    }
                }

                ApplyThemeToChildren(child, bgBrush, fgBrush);
            }
        }

        public void PrepareFormParameters()
        {

            //-------------------------------------------------------------------
            dataTables = GridAccess.GetDataTables();

            gridToHtmlSupported = true;
            foreach (DataTable dt in dataTables)
            {
                if (dt.Rows.Count * dt.Columns.Count > 1000)
                {
                    gridToHtmlSupported = false;
                    break;
                }
            }

            string folderPath = Path.GetTempPath();

            var settings = SettingsManager.GetExcelExportSettings();
            // Suggest a timestamped filename
            string defaultName = ExcelExport.ExpandDateWildcards(settings.GetDefaultFileName());

            exportedFilename = Path.Combine(folderPath, defaultName);

            bool isShiftPressed = Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift);

            ExcelExport.SaveDataTableToExcel(dataTables, exportedFilename, isShiftPressed);

            connectionInfo = ScriptFactoryAccess.GetCurrentConnectionInfo(inMaster: true);
            //\\-----------------------------------------------------------------

            string myEmail = SettingsManager.GetMyEmail();
            CheckBox_CCMyself.IsChecked = true;

            RecipientAddressOptions.Items.Clear();
            RecipientAddressOptions.Items.Add(myEmail);

            List<SettingsManager.FrequentlyUsedEmail> FrequentEmaiss = SettingsManager.GetFrequentlyUsedEmails();
            foreach (var FrequentEmais in FrequentEmaiss)
            {
                if (FrequentEmais.EmailAddress != myEmail)
                    RecipientAddressOptions.Items.Add(FrequentEmais.EmailAddress);
            }

            FullFileNameLabel.Content = exportedFilename;

            FileInfo fileInfo = new FileInfo(exportedFilename);
            double fileSizeInKilobytes = fileInfo.Length / 1024.0;
            string formattedSize = fileSizeInKilobytes.ToString("N0", System.Globalization.CultureInfo.InvariantCulture) + "KB";

            FullFileNameTitleLabel.Content = string.Format(System.Globalization.CultureInfo.CurrentUICulture, "File ({0}):", formattedSize);

            EmailServerOptions.Items.Clear();

         
            //-------------------------------------------------------
            string bodyText =
                "\n\n" + 
                (gridToHtmlSupported ? "{GRID}\n" : "\n") + 
                $"<hr/>\n" +
                $"Data Export\n" +
                $"Server: {connectionInfo.ServerName}\n" +
                $"Database: {connectionInfo.Database}";

            Paragraph para = new Paragraph();
            para.LineHeight = 20; // Set the LineHeight to 20 units
            para.Inlines.Add(new Run(bodyText));
            EmailBody.Document.Blocks.Clear();
            EmailBody.Document.Blocks.Add(para);

            string emailSubject = $"Data Export | Server: {connectionInfo.ServerName} | Database: {connectionInfo.Database}";
            EmailSubject.Text = emailSubject;

            // email and SMTP must be set            
            SettingsManager.SmtpSettings smtpSettings = SettingsManager.GetSmtpSettings();

            if (!string.IsNullOrEmpty(myEmail) && smtpSettings.hasBeenConfiguredAndTested)
            {
                EmailServerOptions.Items.Add(new MailConfigItem(myEmail)); 
            }

            // pull all email profiles from the server:
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionInfo.FullConnectionString))
                {
                    // Open the connection
                    connection.Open();

                    string queryText = @"
                    SELECT sp.[name] AS ProfileName,
                           ma.email_address
                    FROM [msdb].[dbo].[sysmail_profile] AS sp
                         INNER JOIN [msdb].[dbo].[sysmail_profileaccount] AS pa
                            ON sp.[profile_id] = pa.profile_id
                         INNER JOIN [msdb].[dbo].[sysmail_account] AS ma
                            ON pa.[account_id] = ma.[account_id];
                    ";

                    using (SqlCommand command = new SqlCommand(queryText, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                EmailServerOptions.Items.Add(new MailConfigItem(databaseMailProfileName: reader.GetString(0), emailAddress: reader.GetString(1)));
                            }
                        }
                    }      
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error while fetching Database Mail Profiles from the server.");
            }

            if (EmailServerOptions.Items.Count > 0)
            {
                EmailServerOptions.SelectedIndex = 0;
            }
            else
            {
                SendWarningLabel.Content = "Unable to send the email because SMTP wasn't configured and no Database Mail Profiles have been found on the server.";                
                ButtonSend.IsEnabled = false;
            }
        }

        static bool AreAllEmailsValid(string emails)
        {
            string[] emailArray = emails.Split(';');

            foreach (var email in emailArray)
            {
                var trimmedEmail = email.Trim();
                if (!string.IsNullOrEmpty(trimmedEmail))
                {
                    if (!IsValidEmail(trimmedEmail))
                    {
                        // As soon as one invalid email is found, return false
                        return false;
                    }
                }
            }

            // If all emails pass validation, return true
            return true;
        }

        static bool IsValidEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return false;

            try
            {
                var mailAddress = new MailAddress(email);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        static bool SendEmailViaSmtp(MailConfigItem mailConfig, List<String> AllEmails, string ccEmail, string Subject, 
                string Body, string exportedFilename)
        {

            try
            {

                SettingsManager.SmtpSettings smtpSettings = SettingsManager.GetSmtpSettings();

                var smtp = new SmtpClient
                {
                    Host = smtpSettings.ServerName,
                    Port = smtpSettings.Port,
                    EnableSsl = smtpSettings.EnableSsl,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Timeout = 5000 // timeout in milliseconds (e.g., 5 sec)
                };

                if (!string.IsNullOrEmpty(smtpSettings.Username))
                    smtp.Credentials = new NetworkCredential(smtpSettings.Username, smtpSettings.Password);


                var emailMessage = new MailMessage()
                {
                    Subject = Subject,
                    Body = Body,
                    IsBodyHtml = true
                };

                emailMessage.From = new MailAddress(SettingsManager.GetMyEmail());

                foreach (var email in AllEmails)
                    emailMessage.To.Add(email);

                if (!string.IsNullOrEmpty(ccEmail))
                    emailMessage.CC.Add(ccEmail);

                emailMessage.Attachments.Add(new Attachment(exportedFilename));

                smtp.Send(emailMessage);

                MessageBox.Show($"Email has been sent!", "Done");

                try
                {
                    File.Delete(exportedFilename);
                }
                catch (Exception exFile)
                {
                    _logger.Error(exFile, "Error while deleting the exported file after sending the email.");
                }

                return true;

            }
            catch (Exception ex)
            {
                string msg = $"Error message: {ex.Message} \nInnerException: {ex.InnerException}";
                MessageBox.Show(msg, "Something went wrong");
            }

            return false;

        }

        static bool SendEmailViaDatabaseMail(ScriptFactoryAccess.ConnectionInfo connectionInfo, MailConfigItem mailConfig,
            string AllEmails, string ccEmail, string Subject, string Body, string exportedFilename)
        {

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionInfo.FullConnectionString))
                {
                    // Open the connection
                    connection.Open();

                    string queryText = @"
                    USE [msdb];

                    BEGIN TRANSACTION;

                    SET XACT_ABORT ON;

                    DECLARE @mailitem_id AS INT;

                    EXECUTE [dbo].[sp_send_dbmail] 
                        @profile_name       = @ProfileName,
                        --I don't know how service broker works and if it respects open transaction ...
                        @recipients         = @Recipient, 
                        @copy_recipients    = @RecipientCC,
                        @subject            = @Subject, 
                        @body               = @Body, 
                        @body_format        = 'HTML', 
                        @mailitem_id        = @mailitem_id OUTPUT;

                    INSERT INTO sysmail_attachments (mailitem_id, filename, filesize, attachment)
                    SELECT @mailitem_id,
                           @FileName,
                           @FileSize,
                           @BinaryFile;

                    --UPDATE dbo.sysmail_mailitems
                    --SET    recipients = @Recipient
                    --WHERE  mailitem_id = @mailitem_id;

                    COMMIT TRANSACTION;
                    ";

                    using (SqlCommand command = new SqlCommand(queryText, connection))
                    {

                        FileInfo fileInfo = new FileInfo(exportedFilename);
                        byte[] fileContents = File.ReadAllBytes(fileInfo.FullName);

                        command.Parameters.AddWithValue("ProfileName",  mailConfig.MailProfileName);
                        command.Parameters.AddWithValue("Recipient",    AllEmails);
                        command.Parameters.AddWithValue("RecipientCC",  ccEmail);
                        command.Parameters.AddWithValue("Subject",      Subject);
                        command.Parameters.AddWithValue("Body",         Body);
                        command.Parameters.AddWithValue("FileSize",     fileInfo.Length);
                        command.Parameters.AddWithValue("FileName",     fileInfo.Name);
                        command.Parameters.AddWithValue("BinaryFile",   fileContents);

                        command.ExecuteNonQuery();

                    }

                    MessageBox.Show("Email has been queued via Database Mail!", "Done");

                    return true;
                }

            }
            catch (Exception ex)
            {
                string msg = $"Error message: {ex.Message} \nInnerException: {ex.InnerException}";
                MessageBox.Show(msg, "Something went wrong");
            }

            return false;
        }


        private void Button_SendAndClose(object sender, RoutedEventArgs e)
        {

            MailConfigItem mailConfig = (MailConfigItem)EmailServerOptions.SelectedItem;

            bool allValid = AreAllEmailsValid(EmailRecipients.Text);

            if (!allValid)
            {
                MessageBox.Show("Can't parse the recipient's email address.", "Invalid Email");
                return;
            }

            if (string.IsNullOrEmpty(EmailSubject.Text))
            {
                MessageBox.Show("Please provide the email subject.", "Subject Required");
                return;
            }

            string EmailBodyContent = GetEmailBodyHtml();

            List<String> emailClean = new List<string>();
            string[] emailArray = EmailRecipients.Text.Split(';');
            foreach (var email in emailArray)
            {
                var trimmedEmail = email.Trim();
                if (!string.IsNullOrEmpty(trimmedEmail))
                    emailClean.Add(trimmedEmail);
            }
            string allEmailsConcatenated = String.Join(";", emailClean);

            string ccEmail = "";
            if (CheckBox_CCMyself.IsChecked == true)
                ccEmail = SettingsManager.GetMyEmail();

            if (!string.IsNullOrEmpty(TextBoxNewFileName.Text))
            {
                // rename the file
                try
                {
                    string directory = Path.GetDirectoryName(exportedFilename); 
                    string extension = Path.GetExtension(exportedFilename); 
                    string newFileName = TextBoxNewFileName.Text + extension; 
                    string newFilePath = Path.Combine(directory, newFileName); 

                    // Rename the file
                    File.Move(exportedFilename, newFilePath);

                    exportedFilename = newFilePath;

                    FullFileNameLabel.Content = exportedFilename;

                }
                catch {}
            }            

            bool success = false;

            if (mailConfig.isSMTP)
                success = SendEmailViaSmtp(mailConfig, emailClean, ccEmail, EmailSubject.Text, EmailBodyContent, exportedFilename);

            else if (mailConfig.isDatabaseMail)
                success = SendEmailViaDatabaseMail(connectionInfo, mailConfig, allEmailsConcatenated, ccEmail, EmailSubject.Text, EmailBodyContent, exportedFilename);
            
            else {
                MessageBox.Show("Invalid mail config", "Error");
            }

            if (success)
                SettingsManager.SaveFrequentlyUsedEmails(emailClean);


        }

        private void RecipientAddressOptions_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            string email = RecipientAddressOptions.SelectedValue as string;

            if (email != null)
            {
                if (!EmailRecipients.Text.Contains(email))
                {
                    if (!string.IsNullOrEmpty(EmailRecipients.Text))
                        EmailRecipients.Text += "; ";
                    EmailRecipients.Text += email;
                }
            }

        }


    }
}