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

    /// <summary>
    /// Interaction logic for ToolWindowGridToEmailControl.
    /// </summary>
    public partial class ToolWindowGridToEmailControl : UserControl
    {

        public ScriptFactoryAccess.ConnectionInfo connectionInfo;
        public string exportedFilename;
        /// <summary>
        /// Initializes a new instance of the <see cref="ToolWindowGridToEmailControl"/> class.
        /// </summary>
        public ToolWindowGridToEmailControl()
        {
            this.InitializeComponent();

        }

        public void UpdateLabels()
        {
            FullFileName.Text = exportedFilename;

            FileInfo fileInfo = new FileInfo(FullFileName.Text);
            double fileSizeInKilobytes = fileInfo.Length / 1024.0;
            string formattedSize = fileSizeInKilobytes.ToString("N0", System.Globalization.CultureInfo.InvariantCulture) + "KB";

            MyEmailAddress.Content = "From: " + SettingsManager.GetMyEmail();

            FullFileNameLabel.Content = string.Format(System.Globalization.CultureInfo.CurrentUICulture, "File ({0}):", formattedSize);

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
            MessageBox.Show(
                string.Format(System.Globalization.CultureInfo.CurrentUICulture, "Invoked '{0}'", this.ToString()),
                "ToolWindowGridToEmail");
        }

        private void Button_SendAndClose(object sender, RoutedEventArgs e)
        {

            try
            {
                SettingsManager.SmtpSettings smtpSettings = SettingsManager.GetSmtpSettings();

                var fromAddress = new MailAddress(SettingsManager.GetMyEmail());
                var toAddress = new MailAddress(EmailRecipient.Text);

                // Setting up the SMTP client
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

                // Creating the email message
                using (var message = new MailMessage(fromAddress, toAddress)
                {
                    Subject = EmailSubject.Text,
                    Body = EmailBody.Text,

                })
                {

                    Attachment attachment = new Attachment(FullFileName.Text);
                    message.Attachments.Add(attachment);

                    smtp.Send(message);

                    MessageBox.Show(
                        string.Format(System.Globalization.CultureInfo.CurrentUICulture, "Email to '{0}' has been sent!", toAddress),
                        "Done");

                }

                File.Delete(FullFileName.Text);

            } catch (Exception ex)
            {
                string msg = $"Error message: {ex.Message} \nInnerException: {ex.InnerException}";
                MessageBox.Show(msg, "Something went wrong");
            }    
            
        }

        private void Button_SendDbMailAndClose(object sender, RoutedEventArgs e)
        {

            try
            {

                string connectionString = connectionInfo.FullConnectionString;

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    // Open the connection
                    connection.Open();

                    string queryText = @"
                    USE [msdb];

                    BEGIN TRANSACTION;

                    DECLARE @mailitem_id AS INT;

                    EXECUTE [dbo].[sp_send_dbmail] 
                        --I don't know how service broker works and it respects open transaction
                        --I want to avoid a case when email will be sent before attachment is inserted
                        @recipients     = '< email placeholder >', 
                        @subject        = @Subject, 
                        @body           = @Body, 
                        @body_format    = 'HTML', 
                        @mailitem_id    = @mailitem_id OUTPUT;

                    INSERT INTO sysmail_attachments (mailitem_id, filename, filesize, attachment)
                    SELECT @mailitem_id,
                           @FileName,
                           @FileSize,
                           @BinaryFile;

                    UPDATE dbo.sysmail_mailitems
                    SET    recipients = @Recipient
                    WHERE  mailitem_id = @mailitem_id;

                    COMMIT TRANSACTION;
                    ";

                    using (SqlCommand command = new SqlCommand(queryText, connection))
                    {

                        FileInfo fileInfo = new FileInfo(exportedFilename);
                        byte[] fileContents = File.ReadAllBytes(fileInfo.FullName);

                        command.Parameters.AddWithValue("Recipient",    EmailRecipient.Text);
                        command.Parameters.AddWithValue("Subject",      EmailSubject.Text);
                        command.Parameters.AddWithValue("Body",         EmailBody.Text);
                        command.Parameters.AddWithValue("FileSize",     fileInfo.Length);
                        command.Parameters.AddWithValue("FileName",     fileInfo.Name);
                        command.Parameters.AddWithValue("BinaryFile",   fileContents);

                        command.ExecuteNonQuery();

                    }

                }

            }
            catch (Exception ex)
            {
                string msg = $"Error message: {ex.Message} \nInnerException: {ex.InnerException}";
                MessageBox.Show(msg, "Something went wrong");
            }
        }

    }
}