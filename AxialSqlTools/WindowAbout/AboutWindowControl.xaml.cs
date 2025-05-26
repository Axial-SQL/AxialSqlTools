namespace AxialSqlTools
{
    using System;
    using System.Diagnostics;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.Reflection;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Navigation;

    /// <summary>
    /// Interaction logic for AboutWindowControl.
    /// </summary>
    public partial class AboutWindowControl : UserControl
    {

        private string _logFolder;
        /// <summary>
        /// Initializes a new instance of the <see cref="AboutWindowControl"/> class.
        /// </summary>
        public AboutWindowControl()
        {
            this.InitializeComponent();

            Version currentVersion = Assembly.GetExecutingAssembly().GetName().Version;
            string currentVersionString = currentVersion.ToString();

            TextBlock_CurrentVersion.Text = $"Axial SQL Tools | SSMS Addin Version {currentVersionString}";

            _logFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "AxialSQL",
                        "AxialSQLToolsLog"
                );

            HyperlinkText_LogFolder.Text = _logFolder;
        }

        private void Hyperlink_RequestNavigateEmail(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
        }

        private void buttonAxialSqlWebsite_Click(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
        }


        private void HyperlinkLogFolder_Click(object sender, RoutedEventArgs e)
        {

            if (!Directory.Exists(_logFolder))
            {
                Directory.CreateDirectory(_logFolder);
            }

            // Open the folder in Windows Explorer
            Process.Start("explorer.exe", _logFolder);
        }
    }
}