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
        }

        public void PrepareFormParameters()
        {

            //-------------------------------------------------------------------
            dataTables = GridAccess.GetDataTables();

            gridToHtmlSupported = true;
            foreach (DataTable dt in dataTables)
            {
                if (dt.Rows.Count * dt.Columns.Count < 1000)
                {
                    gridToHtmlSupported = false;
                    break;
                }
            }


            string folderPath = Path.GetTempPath();
            string fileName = $"DataExport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            exportedFilename = Path.Combine(folderPath, fileName);

            ExcelExport.SaveDataTableToExcel(dataTables, exportedFilename);

            connectionInfo = ScriptFactoryAccess.GetCurrentConnectionInfo(inMaster: true);
            //\\-----------------------------------------------------------------

            string myEmail = SettingsManager.GetMyEmail();

            RecipientAddressOptions.Items.Clear();
            RecipientAddressOptions.Items.Add(myEmail);
            //TODO - persist/restore common used emails

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

        static void SendEmailViaSmtp(MailConfigItem mailConfig, List<String> AllEmails, string Subject, 
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
                    UseDefaultCredentials = false
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

                emailMessage.Attachments.Add(new Attachment(exportedFilename));

                smtp.Send(emailMessage);

                MessageBox.Show($"Email has been sent!", "Done");

                try
                {
                    File.Delete(exportedFilename);
                }
                catch (Exception exFile)
                { }

            }
            catch (Exception ex)
            {
                string msg = $"Error message: {ex.Message} \nInnerException: {ex.InnerException}";
                MessageBox.Show(msg, "Something went wrong");
            }

        }

        static void SendEmailViaDatabaseMail(ScriptFactoryAccess.ConnectionInfo connectionInfo, MailConfigItem mailConfig,
            string AllEmails, string Subject, string Body, string exportedFilename)
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
                        @profile_name = @ProfileName,
                        --I don't know how service broker works and if it respects open transaction ...
                        @recipients     = @Recipient, 
                        @subject        = @Subject, 
                        @body           = @Body, 
                        @body_format    = 'HTML', 
                        @mailitem_id    = @mailitem_id OUTPUT;

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
                        command.Parameters.AddWithValue("Subject",      Subject);
                        command.Parameters.AddWithValue("Body",         Body);
                        command.Parameters.AddWithValue("FileSize",     fileInfo.Length);
                        command.Parameters.AddWithValue("FileName",     fileInfo.Name);
                        command.Parameters.AddWithValue("BinaryFile",   fileContents);

                        command.ExecuteNonQuery();

                    }

                    MessageBox.Show("Email has been queued via Database Mail!", "Done");
                }

            }
            catch (Exception ex)
            {
                string msg = $"Error message: {ex.Message} \nInnerException: {ex.InnerException}";
                MessageBox.Show(msg, "Something went wrong");
            }
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


            if (mailConfig.isSMTP)
                SendEmailViaSmtp(mailConfig, emailClean, EmailSubject.Text, EmailBodyContent, exportedFilename);

            else if (mailConfig.isDatabaseMail)
                SendEmailViaDatabaseMail(connectionInfo, mailConfig, allEmailsConcatenated, EmailSubject.Text, EmailBodyContent, exportedFilename);
            
            else {
                MessageBox.Show("Invalid mail config", "Error");
            }

            
        }

        private void RecipientAddressOptions_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            string email = RecipientAddressOptions.SelectedValue as string;

            if (!EmailRecipients.Text.Contains(email))
            {
                if (!string.IsNullOrEmpty(EmailRecipients.Text))
                    EmailRecipients.Text += "; ";
                EmailRecipients.Text += email;
            }

        }


    }
}