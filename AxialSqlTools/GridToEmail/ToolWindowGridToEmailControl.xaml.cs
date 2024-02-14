namespace AxialSqlTools
{
    using Microsoft.VisualStudio.Shell;
    using System;
    using System.Data.SqlClient;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.Net;
    using System.Net.Mail;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Documents;

    /// <summary>
    /// Interaction logic for ToolWindowGridToEmailControl.
    /// </summary>
    public partial class ToolWindowGridToEmailControl : UserControl
    {

        public ScriptFactoryAccess.ConnectionInfo connectionInfo;
        public string exportedFilename;

        public class MailConfigItem
        {
            //public int Id { get; set; }
            //public string Name { get; }

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
            FullFileNameLabel.Content = exportedFilename;

            FileInfo fileInfo = new FileInfo(exportedFilename);
            double fileSizeInKilobytes = fileInfo.Length / 1024.0;
            string formattedSize = fileSizeInKilobytes.ToString("N0", System.Globalization.CultureInfo.InvariantCulture) + "KB";

            FullFileNameTitleLabel.Content = string.Format(System.Globalization.CultureInfo.CurrentUICulture, "File ({0}):", formattedSize);

            EmailServerOptions.Items.Clear();

         
            //-------------------------------------------------------
            string bodyText =
                $"\n\n\n--------------------------\n" +
                $"Data Export\n" +
                $"Server: {connectionInfo.ServerName}\n" +
                $"Database: {connectionInfo.Database}";

            Paragraph para = new Paragraph();
            para.LineHeight = 20; // Set the LineHeight to 20 units
            para.Inlines.Add(new Run(bodyText));
            EmailBody.Document.Blocks.Clear();
            EmailBody.Document.Blocks.Add(para);


            // email and SMTP must be set
            string myEmail = SettingsManager.GetMyEmail();
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
                if (!IsValidEmail(email))
                {
                    // As soon as one invalid email is found, return false
                    return false;
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

        static void SendEmailViaSmtp(MailConfigItem mailConfig, string AllEmails, string Subject, 
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
                    Body = Body
                };

                emailMessage.From = new MailAddress(SettingsManager.GetMyEmail());


                string[] emailArray = AllEmails.Split(';');
                foreach (var email in emailArray)
                    emailMessage.To.Add(email);

                emailMessage.Attachments.Add(new Attachment(exportedFilename));

                smtp.Send(emailMessage);

                MessageBox.Show($"Email to '{AllEmails}' has been sent!", "Done");

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

            TextRange textRange = new TextRange(EmailBody.Document.ContentStart, EmailBody.Document.ContentEnd);
            string EmailBodyContent = textRange.Text;

            if (mailConfig.isSMTP)
                SendEmailViaSmtp(mailConfig, EmailRecipients.Text, EmailSubject.Text, EmailBodyContent, exportedFilename);

            else if (mailConfig.isDatabaseMail)
                SendEmailViaDatabaseMail(connectionInfo, mailConfig, EmailRecipients.Text, EmailSubject.Text, EmailBodyContent, exportedFilename);
            
            else {
                MessageBox.Show("Invalid mail config", "Error");
            }

            
        }



    }
}