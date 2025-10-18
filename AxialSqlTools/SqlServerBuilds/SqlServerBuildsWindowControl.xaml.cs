using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using static AxialSqlTools.AxialSqlToolsPackage;

namespace AxialSqlTools
{
    /// <summary>
    /// Interaction logic for SqlServerBuildsWindowControl.
    /// </summary>
    public partial class SqlServerBuildsWindowControl : UserControl
    {
        public SqlServerBuildsWindowControl()
        {
            this.InitializeComponent();
            LoadSqlVersions();
        }

        private void LoadSqlVersions()
        {
            var sqlVersions = AxialSqlToolsPackage.PackageInstance.SQLBuildsDataInfo;
            sqlVersionTreeView.Items.Clear(); // Clear previous items

            // Add Group Headers
            foreach (var majorVersion in sqlVersions.Builds)
            {
                TreeViewItem groupNode = new TreeViewItem
                {
                    Header = $"SQL Server {majorVersion.Key}",
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.DarkBlue,
                    IsExpanded = true
                };

                // Add Column Headers (Styled like a ListView)
                StackPanel headerPanel = new StackPanel { Orientation = Orientation.Horizontal, Background = Brushes.LightGray };
                headerPanel.Children.Add(CreateColumnText("Build Number", 120, FontWeights.Bold));
                headerPanel.Children.Add(CreateColumnText("KB Number", 120, FontWeights.Bold));
                headerPanel.Children.Add(CreateColumnText("Release Date", 120, FontWeights.Bold));
                headerPanel.Children.Add(CreateColumnText("Update Name", 100, FontWeights.Bold));
                headerPanel.Children.Add(CreateColumnText("URL", 150, FontWeights.Bold));

                groupNode.Items.Add(new TreeViewItem { Header = headerPanel, IsEnabled = false });

                // Add Child Rows
                foreach (var update in majorVersion.Value)
                {
                    StackPanel rowPanel = new StackPanel { Orientation = Orientation.Horizontal };

                    if (update.BuildNumber != null)
                    {
                        rowPanel.Children.Add(CreateColumnText(update.BuildNumber.ToString(), 120));
                    }
                    else
                    {
                        rowPanel.Children.Add(CreateColumnText("n/a", 120));
                    }
                    rowPanel.Children.Add(CreateColumnText(update.KbNumber, 100));
                    rowPanel.Children.Add(CreateColumnText(update.ReleaseDate.ToString("yyyy-MM-dd"), 120));
                    rowPanel.Children.Add(CreateColumnText(update.UpdateName, 100));

                    TextBlock link = CreateHyperlink(update.Url, 150);
                    rowPanel.Children.Add(link);

                    TreeViewItem updateNode = new TreeViewItem { Header = rowPanel };
                    groupNode.Items.Add(updateNode);
                }

                sqlVersionTreeView.Items.Add(groupNode);
            }
        }

        // Helper to create text column
        private TextBlock CreateColumnText(string text, double width, FontWeight fontWeight = default)
        {
            return new TextBlock
            {
                Text = text,
                Width = width,
                FontWeight = fontWeight,
                Margin = new Thickness(5, 2, 5, 2)
            };
        }

        // Helper to create hyperlink column
        private TextBlock CreateHyperlink(string url, double width)
        {
            TextBlock link = new TextBlock
            {
                Text = "More Info",
                Width = width,
                Foreground = Brushes.Blue,
                TextDecorations = TextDecorations.Underline,
                Cursor = System.Windows.Input.Cursors.Hand
            };
            link.MouseLeftButtonUp += (s, e) => OpenUrl(url);
            return link;
        }

        private static string ExtractUrlAfter_HYPERLINK(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return null;

            int lastParen = input.LastIndexOf(')');
            if (lastParen >= 0 && lastParen < input.Length - 1)
            {
                return input.Substring(lastParen + 1).Trim();
            }

            return input.Trim();
        }

        private void OpenUrl(string url)
        {
            try
            {
                if (url.StartsWith("HYPERLINK("))
                {
                    url = ExtractUrlAfter_HYPERLINK(url);
                }
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });

            } catch (Exception ex)
            {
                VsShellUtilities.ShowMessageBox(
                    ServiceProvider.GlobalProvider,
                    $"Failed to open URL:\n{url}\n\nError: {ex.Message}",
                    "Open URL Error",
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        private void HyperlinkDataSource_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(new ProcessStartInfo("https://aka.ms/sqlserverbuilds") { UseShellExecute = true });
        }

        /// <summary>
        /// Event handler for the "Copy as TSQL" hyperlink.
        /// Generates a T-SQL script that creates a temp table and populates it with the SQLBuildsData,
        /// then copies the script to the clipboard.
        /// </summary>
        private void HyperlinkCopyAsTSQL_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var sqlData = AxialSqlToolsPackage.PackageInstance.SQLBuildsDataInfo;
            StringBuilder sb = new StringBuilder();

            // Create the temporary table script
            sb.AppendLine("-- Script generated by AxialSqlTools");
            sb.AppendLine("CREATE TABLE #SQLBuildsData (");
            sb.AppendLine("    [MajorVersion] NVARCHAR(50),");
            sb.AppendLine("    [BuildNumber]  NVARCHAR(50),");
            sb.AppendLine("    [KbNumber]     NVARCHAR(50),");
            sb.AppendLine("    [ReleaseDate]  DATE,");
            sb.AppendLine("    [UpdateName]   NVARCHAR(100),");
            sb.AppendLine("    [Url]          NVARCHAR(255)");
            sb.AppendLine(");");
            sb.AppendLine();

            // Loop through the builds data and generate INSERT statements
            foreach (var majorVersion in sqlData.Builds)
            {
                foreach (var update in majorVersion.Value)
                {
                    // Prepare values with basic SQL escaping for single quotes
                    string major = majorVersion.Key.Replace("'", "''");
                    string buildNumber = update.BuildNumber != null ? update.BuildNumber.ToString() : "n/a";
                    buildNumber = buildNumber.Replace("'", "''");
                    string releaseDate = update.ReleaseDate.ToString("yyyy-MM-dd");
                    string kbNumber = update.KbNumber ?? "N/A";
                    string updateName = update.UpdateName != null ? update.UpdateName.Replace("'", "''") : "";
                    string url = update.Url != null ? update.Url.Replace("'", "''") : "";
                    if (url.StartsWith("HYPERLINK("))
                    {
                        url = ExtractUrlAfter_HYPERLINK(url);
                    }

                    sb.AppendLine($"INSERT INTO #SQLBuildsData ([MajorVersion], [BuildNumber], [KbNumber], [ReleaseDate], [UpdateName], [Url]) VALUES ('{major}', '{buildNumber}', '{kbNumber}', '{releaseDate}', '{updateName}', '{url}');");
                }
            }

            sb.AppendLine();
            sb.AppendLine("SELECT * FROM #SQLBuildsData;");
            sb.AppendLine();

            // Copy the generated script to the clipboard
            Clipboard.SetDataObject(sb.ToString());
            MessageBox.Show("TSQL script copied to clipboard!", "Copy as TSQL", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
