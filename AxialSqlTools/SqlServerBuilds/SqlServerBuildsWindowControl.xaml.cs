namespace AxialSqlTools
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Diagnostics;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Documents;
    using System.Windows.Media;
    using static AxialSqlTools.AxialSqlToolsPackage;

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
                headerPanel.Children.Add(CreateColumnText("Release Date", 120, FontWeights.Bold));
                headerPanel.Children.Add(CreateColumnText("Update Name", 100, FontWeights.Bold));
                headerPanel.Children.Add(CreateColumnText("URL", 150, FontWeights.Bold));

                groupNode.Items.Add(new TreeViewItem { Header = headerPanel, IsEnabled = false });

                // Add Child Rows
                foreach (var update in majorVersion.Value)
                {
                    StackPanel rowPanel = new StackPanel { Orientation = Orientation.Horizontal };

                    if (!(update.BuildNumber == null))
                    {
                        rowPanel.Children.Add(CreateColumnText(update.BuildNumber.ToString(), 120));
                    }
                    else {
                        rowPanel.Children.Add(CreateColumnText("n/a", 120));
                    }
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
            link.MouseLeftButtonUp += (s, e) => Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            return link;
        }

        private void HyperlinkDataSource_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(new ProcessStartInfo("https://aka.ms/sqlserverbuilds") { UseShellExecute = true });
        }

    }
}