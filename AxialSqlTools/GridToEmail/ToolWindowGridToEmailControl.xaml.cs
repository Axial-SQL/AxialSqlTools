namespace AxialSqlTools
{
    using Microsoft.VisualStudio.Shell;
    using System;
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
        /// <summary>
        /// Initializes a new instance of the <see cref="ToolWindowGridToEmailControl"/> class.
        /// </summary>
        public ToolWindowGridToEmailControl()
        {
            this.InitializeComponent();

        }

        public void UpdateLabels()
        {

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
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(smtpSettings.Username, smtpSettings.Password)
                };

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
                string msg = $"Erorr message: {ex.Message} \nInnerException: {ex.InnerException}";
                MessageBox.Show(msg, "Something went wrong");
            }    
            
        }
    }
}