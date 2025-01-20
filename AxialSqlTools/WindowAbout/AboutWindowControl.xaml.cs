namespace AxialSqlTools
{
    using System;
    using System.IO;
    using System.Diagnostics;
    using System.Diagnostics.CodeAnalysis;
    using System.Reflection;
    using System.Windows;
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for AboutWindowControl.
    /// </summary>
    public partial class AboutWindowControl : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AboutWindowControl"/> class.
        /// </summary>
        public AboutWindowControl()
        {
            this.InitializeComponent();

            Version currentVersion = Assembly.GetExecutingAssembly().GetName().Version;
            string currentVersionString = currentVersion.ToString();

            TextBlock_CurrentVersion.Text = $"Axial SQL Tools | SSMS Addin Version {currentVersionString}";
        }

        /// <summary>
        /// Handles click on the button by displaying a message box.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event args.</param>
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Justification = "Sample code")]
        [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1300:ElementMustBeginWithUpperCaseLetter", Justification = "Default event handler naming pattern")]
        private void buttonAxialSqlWebsite_Click(object sender, RoutedEventArgs e)
        {

            string url = "https://axial-sql.com";
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });

        }

        private void ButtonOpenLogFolder_Click(object sender, RoutedEventArgs e)
        {


            var logDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "AxialSQL",
                        "AxialSQLToolsLog"
                );

            if (!Directory.Exists(logDirectory))
            {
                Directory.CreateDirectory(logDirectory);
            }

            // Open the folder in Windows Explorer
            Process.Start("explorer.exe", logDirectory);
        }
    }
}