using System;
using System.Diagnostics;
using System.IO;
using System.Windows;

namespace AxialSqlTools
{
    public partial class ExcelExportSuccessDialog : Window
    {
        private readonly ToolWindowThemeController _themeController;
        private readonly string filePath;

        public ExcelExportSuccessDialog(string filePath)
        {
            InitializeComponent();

            _themeController = new ToolWindowThemeController(this, ApplyThemeBrushResources);

            this.filePath = filePath ?? string.Empty;
            FilePathText.Text = this.filePath;
        }

        private void ApplyThemeBrushResources()
        {
            ToolWindowThemeResources.ApplySharedTheme(this);
        }

        private void OpenInExcelButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show(this, "The exported file could not be found.", "Open in Excel", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                ProcessStartInfo processStartInfo = new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                };

                Process.Start(processStartInfo);
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Unable to open the exported file in Excel.\n\n" + ex.Message, "Open in Excel", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
