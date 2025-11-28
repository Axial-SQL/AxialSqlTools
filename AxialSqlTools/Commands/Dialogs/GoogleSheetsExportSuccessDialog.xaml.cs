using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Navigation;

namespace AxialSqlTools
{
    public partial class GoogleSheetsExportSuccessDialog : Window
    {
        private readonly string spreadsheetUrl;

        public GoogleSheetsExportSuccessDialog(string spreadsheetUrl, string spreadsheetTitle)
        {
            InitializeComponent();

            this.spreadsheetUrl = spreadsheetUrl ?? string.Empty;
            SpreadsheetLinkText.Text = string.IsNullOrWhiteSpace(spreadsheetTitle) ? spreadsheetUrl : spreadsheetTitle;

            Uri uri;
            if (Uri.TryCreate(spreadsheetUrl, UriKind.Absolute, out uri))
            {
                SpreadsheetLink.NavigateUri = uri;
            }
        }

        private void SpreadsheetLink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(spreadsheetUrl))
            {
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = spreadsheetUrl,
                        UseShellExecute = true
                    };
                    Process.Start(psi);
                }
                catch
                {
                    // Swallow exceptions from launching the browser to keep the dialog responsive.
                }
            }

            e.Handled = true;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
