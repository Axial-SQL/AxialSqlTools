using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
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
        private readonly ToolWindowThemeController _themeController;

        public SqlServerBuildsWindowControl()
        {
            this.InitializeComponent();
            _themeController = new ToolWindowThemeController(this, ApplyThemeBrushResources);
            LoadSqlVersions();
        }

        private void ApplyThemeBrushResources()
        {
            ToolWindowThemeResources.ApplySharedTheme(this);
            LoadSqlVersions();
        }

        private void LoadSqlVersions()
        {
            var sqlVersions = AxialSqlToolsPackage.PackageInstance.SQLBuildsDataInfo;
            sqlVersionTreeView.Items.Clear(); // Clear previous items

            Brush groupHeaderBrush = ResolveThemeBrush("AxialThemeAccentBrush", Brushes.DarkBlue);
            Brush headerBackgroundBrush = ResolveThemeBrush("AxialThemeGridHeaderBackgroundBrush", Brushes.LightGray);
            Brush foregroundBrush = ResolveThemeBrush("AxialThemeForegroundBrush", Brushes.Black);
            Brush linkBrush = ResolveThemeBrush("AxialThemeLinkBrush", Brushes.Blue);

            // Add Group Headers
            foreach (var majorVersion in sqlVersions.Builds)
            {
                TreeViewItem groupNode = new TreeViewItem
                {
                    Header = $"SQL Server {majorVersion.Key}",
                    FontWeight = FontWeights.Bold,
                    Foreground = groupHeaderBrush,
                    IsExpanded = true
                };

                // Add Column Headers (Styled like a ListView)
                StackPanel headerPanel = new StackPanel { Orientation = Orientation.Horizontal, Background = headerBackgroundBrush };
                headerPanel.Children.Add(CreateColumnText("Build Number", 120, foregroundBrush, FontWeights.Bold));
                headerPanel.Children.Add(CreateColumnText("KB Number", 120, foregroundBrush, FontWeights.Bold));
                headerPanel.Children.Add(CreateColumnText("Release Date", 120, foregroundBrush, FontWeights.Bold));
                headerPanel.Children.Add(CreateColumnText("Update Name", 100, foregroundBrush, FontWeights.Bold));
                headerPanel.Children.Add(CreateColumnText("URL", 150, foregroundBrush, FontWeights.Bold));

                groupNode.Items.Add(new TreeViewItem { Header = headerPanel, IsEnabled = false });

                // Add Child Rows
                foreach (var update in majorVersion.Value)
                {
                    StackPanel rowPanel = new StackPanel { Orientation = Orientation.Horizontal };

                    if (update.BuildNumber != null)
                    {
                        rowPanel.Children.Add(CreateColumnText(update.BuildNumber.ToString(), 120, foregroundBrush));
                    }
                    else
                    {
                        rowPanel.Children.Add(CreateColumnText("n/a", 120, foregroundBrush));
                    }
                    rowPanel.Children.Add(CreateColumnText(update.KbNumber, 100, foregroundBrush));
                    rowPanel.Children.Add(CreateColumnText(update.ReleaseDate.ToString("yyyy-MM-dd"), 120, foregroundBrush));
                    rowPanel.Children.Add(CreateColumnText(update.UpdateName, 100, foregroundBrush));

                    TextBlock link = CreateHyperlink(update.Url, 150, linkBrush);
                    rowPanel.Children.Add(link);

                    TreeViewItem updateNode = new TreeViewItem { Header = rowPanel };
                    groupNode.Items.Add(updateNode);
                }

                sqlVersionTreeView.Items.Add(groupNode);
            }
        }

        // Helper to create text column
        private TextBlock CreateColumnText(string text, double width, Brush foregroundBrush, FontWeight fontWeight = default)
        {
            return new TextBlock
            {
                Text = text,
                Width = width,
                Foreground = foregroundBrush,
                FontWeight = fontWeight,
                Margin = new Thickness(5, 2, 5, 2)
            };
        }

        // Helper to create hyperlink column
        private TextBlock CreateHyperlink(string url, double width, Brush foregroundBrush)
        {
            string normalizedUrl = NormalizeUrl(url);
            TextBlock link = new TextBlock
            {
                Text = "More Info",
                Width = width,
                Foreground = foregroundBrush,
                TextDecorations = TextDecorations.Underline,
                Cursor = System.Windows.Input.Cursors.Hand
            };

            MenuItem copyUrlMenuItem = new MenuItem
            {
                Header = "Copy URL"
            };
            copyUrlMenuItem.Click += (s, e) => CopyUrlToClipboard(normalizedUrl);

            ContextMenu contextMenu = new ContextMenu();
            contextMenu.Items.Add(copyUrlMenuItem);
            link.ContextMenu = contextMenu;

            link.MouseLeftButtonUp += (s, e) => OpenUrl(normalizedUrl);
            return link;
        }

        private void CopyUrlToClipboard(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                VsShellUtilities.ShowMessageBox(
                    ServiceProvider.GlobalProvider,
                    "No URL is available to copy.",
                    "Copy URL",
                    OLEMSGICON.OLEMSGICON_WARNING,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                return;
            }

            Clipboard.SetDataObject(url);
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

        private static string NormalizeUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
                return url;

            return url.StartsWith("HYPERLINK(") ? ExtractUrlAfter_HYPERLINK(url) : url;
        }

        private Brush ResolveThemeBrush(string resourceKey, Brush fallback)
        {
            return TryFindResource(resourceKey) as Brush ?? fallback;
        }

        private void OpenUrl(string url)
        {
            try
            {
                url = NormalizeUrl(url);
                if (!ToolWindowNavigation.OpenExternalUrl(url))
                {
                    throw new InvalidOperationException("URL is empty.");
                }

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
            OpenUrl("https://aka.ms/sqlserverbuilds");
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
            sb.AppendLine("-- Source https://aka.ms/sqlserverbuilds");
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
                    url = NormalizeUrl(url);

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
