using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Win32;
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
            Loaded += ControlLoaded;
            Unloaded += ControlUnloaded;
            LoadSqlVersions();
        }

        private AxialSqlToolsPackage subscribedPackage;

        private void ControlLoaded(object sender, RoutedEventArgs e)
        {
            ControlUnloaded(sender, e);
            subscribedPackage = AxialSqlToolsPackage.PackageInstance;
            if (subscribedPackage != null) subscribedPackage.SQLBuildsChanged += BuildsChanged;
            LoadSqlVersions();
        }

        private void ControlUnloaded(object sender, RoutedEventArgs e)
        {
            if (subscribedPackage != null) subscribedPackage.SQLBuildsChanged -= BuildsChanged;
            subscribedPackage = null;
        }

        private void BuildsChanged(object sender, EventArgs e)
        {
            if (!Dispatcher.CheckAccess()) { Dispatcher.BeginInvoke(new Action(LoadSqlVersions)); return; }
            LoadSqlVersions();
        }

        private async void Refresh_Click(object sender, RoutedEventArgs e) => await RefreshAsync(null);
        private async void OpenWorkbook_Click(object sender, RoutedEventArgs e)
        {
            var picker = new OpenFileDialog { Filter = "Excel workbook (*.xlsx)|*.xlsx", CheckFileExists = true };
            if (picker.ShowDialog(Window.GetWindow(this)) == true) await RefreshAsync(picker.FileName);
        }

        private async Task RefreshAsync(string path)
        {
            try
            {
                var package = AxialSqlToolsPackage.PackageInstance;
                if (package == null) throw new InvalidOperationException("The add-in is still initializing. Try again shortly.");
                await package.RefreshSqlServerBuildsAsync(path);
            }
            catch (Exception ex) { ShowActionError(ex); }
        }

        private void ShowActionError(Exception ex)
        {
            StatusText.Text = ex.Message;
            StatusText.Foreground = ResolveThemeBrush("AxialThemeStatusErrorBrush", Brushes.Firebrick);
        }

        private void ShowLoadStatus(AxialSqlToolsPackage package)
        {
            var state = package?.SQLBuildsLoadState;
            var success = package?.SQLBuildsLastSuccess;
            bool hasData = package?.SQLBuildsDataInfo?.Builds?.Values.Any(v => v != null && v.Count > 0) == true;
            bool loading = package?.SQLBuildsIsLoading == true;
            RefreshButton.IsEnabled = OpenWorkbookButton.IsEnabled = !loading;
            CopySqlLink.IsEnabled = hasData;
            LoadingProgress.Visibility = loading ? Visibility.Visible : Visibility.Collapsed;
            RefreshButton.Content = state?.Error != null ? "Retry download" : "Refresh";
            string retained = hasData && success != null ? " Showing previously loaded data from " + success.LoadedUtc.ToLocalTime().ToString("g") + "." : " No build data is available.";
            StatusText.Text = loading ? "Loading SQL Server builds..." + (hasData ? retained : "") :
                state?.Error != null ? state.Error + retained :
                state == null ? "Build information has not loaded yet. Select Refresh to retry." :
                state.Warnings.Count > 0 ? "Workbook loaded with issues. See details for missing worksheets, skipped rows, or unavailable values." : "Build information loaded successfully.";
            StatusText.Foreground = state?.Error != null && !loading ? ResolveThemeBrush("AxialThemeStatusErrorBrush", Brushes.Firebrick) :
                ResolveThemeBrush("AxialThemeForegroundBrush", Brushes.Black);
            SourceText.Text = success == null ? "Source: " + SQLBuilds.SourceUrl : "Displayed data: " + success.Source + " | Loaded: " + success.LoadedUtc.ToLocalTime().ToString("g");
            string details = state == null ? "" : string.Join(Environment.NewLine, new[] { state.Error, state.Details }.Where(v => !string.IsNullOrWhiteSpace(v)).Concat(state.Warnings));
            if (state != null && state.Error != null) details = "Attempted source: " + state.Source + Environment.NewLine + details;
            DiagnosticText.Text = details;
            Diagnostics.Visibility = string.IsNullOrWhiteSpace(details) ? Visibility.Collapsed : Visibility.Visible;
            EmptyText.Visibility = hasData ? Visibility.Collapsed : Visibility.Visible;
        }

        private void ApplyThemeBrushResources()
        {
            ToolWindowThemeResources.ApplySharedTheme(this);
            LoadSqlVersions();
        }

        private void LoadSqlVersions()
        {
            var package = AxialSqlToolsPackage.PackageInstance;
            ShowLoadStatus(package);
            var sqlVersions = package?.SQLBuildsDataInfo;
            sqlVersionTreeView.Items.Clear(); // Clear previous items
            if (sqlVersions?.Builds == null) return;

            Brush groupHeaderBrush = ResolveThemeBrush("AxialThemeAccentBrush", Brushes.DarkBlue);
            Brush headerBackgroundBrush = ResolveThemeBrush("AxialThemeGridHeaderBackgroundBrush", Brushes.LightGray);
            Brush foregroundBrush = ResolveThemeBrush("AxialThemeForegroundBrush", Brushes.Black);
            Brush linkBrush = ResolveThemeBrush("AxialThemeLinkBrush", Brushes.Blue);

            // Add Group Headers
            foreach (var majorVersion in sqlVersions.Builds.OrderByDescending(v => v.Key))
            {
                TreeViewItem groupNode = new TreeViewItem
                {
                    Header = $"SQL Server {majorVersion.Key}",
                    FontWeight = FontWeights.Bold,
                    Foreground = groupHeaderBrush,
                    IsExpanded = true
                };

                if (majorVersion.Value == null) continue;

                // Add Column Headers (Styled like a ListView)
                StackPanel headerPanel = new StackPanel { Orientation = Orientation.Horizontal, Background = headerBackgroundBrush };
                headerPanel.Children.Add(CreateColumnText("Build Number", 120, foregroundBrush, FontWeights.Bold));
                headerPanel.Children.Add(CreateColumnText("KB Number", 100, foregroundBrush, FontWeights.Bold));
                headerPanel.Children.Add(CreateColumnText("Release Date", 120, foregroundBrush, FontWeights.Bold));
                headerPanel.Children.Add(CreateColumnText("Update Name", 320, foregroundBrush, FontWeights.Bold));
                headerPanel.Children.Add(CreateColumnText("URL", 150, foregroundBrush, FontWeights.Bold));

                groupNode.Items.Add(new TreeViewItem { Header = headerPanel, IsEnabled = false });

                // Add Child Rows
                foreach (var update in majorVersion.Value.Where(v => v != null))
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
                    rowPanel.Children.Add(CreateColumnText((update.ReleaseDate == default(DateTime) ? "Unknown" : update.ReleaseDate.ToString("yyyy-MM-dd")), 120, foregroundBrush));
                    rowPanel.Children.Add(CreateColumnText(update.UpdateName, 320, foregroundBrush));

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
                Text = text ?? "",
                ToolTip = text,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Width = width,
                Foreground = foregroundBrush,
                FontWeight = fontWeight,
                Margin = new Thickness(5, 2, 5, 2)
            };
        }

        // Helper to create hyperlink column
        private TextBlock CreateHyperlink(string url, double width, Brush foregroundBrush)
        {
            if (!SQLBuilds.TryHttpUrl(url, out string normalizedUrl)) return CreateColumnText("Unavailable", width, foregroundBrush);
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

            try { Clipboard.SetDataObject(url); } catch (Exception ex) { ShowActionError(ex); }
        }

        private Brush ResolveThemeBrush(string resourceKey, Brush fallback)
        {
            return TryFindResource(resourceKey) as Brush ?? fallback;
        }

        private void OpenUrl(string url)
        {
            try
            {
                if (!SQLBuilds.TryHttpUrl(url, out string validUrl)) throw new InvalidOperationException("Only valid HTTP or HTTPS links can be opened.");
                url = validUrl;
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

            var sqlData = AxialSqlToolsPackage.PackageInstance?.SQLBuildsDataInfo;
            if (sqlData?.Builds == null || !sqlData.Builds.Values.Any(v => v != null && v.Count > 0))
            {
                StatusText.Text = "There is no build data to copy. Retry the download or open a readable .xlsx workbook.";
                return;
            }
            StringBuilder sb = new StringBuilder();

            // Create the temporary table script
            sb.AppendLine("-- Source https://aka.ms/sqlserverbuilds");
            sb.AppendLine("CREATE TABLE #SQLBuildsData (");
            sb.AppendLine("    [MajorVersion] NVARCHAR(50),");
            sb.AppendLine("    [BuildNumber]  NVARCHAR(50),");
            sb.AppendLine("    [KbNumber]     NVARCHAR(50),");
            sb.AppendLine("    [ReleaseDate]  DATE,");
            sb.AppendLine("    [UpdateName]   NVARCHAR(MAX),");
            sb.AppendLine("    [Url]          NVARCHAR(MAX)");
            sb.AppendLine(");");
            sb.AppendLine();

            // Loop through the builds data and generate INSERT statements
            foreach (var majorVersion in sqlData.Builds)
            {
                foreach (var update in majorVersion.Value ?? new List<SQLVersionInfo>())
                {
                    if (update == null) continue;
                    // Prepare values with basic SQL escaping for single quotes
                    string major = majorVersion.Key.Replace("'", "''");
                    string buildNumber = update.BuildNumber != null ? update.BuildNumber.ToString() : "n/a";
                    buildNumber = buildNumber.Replace("'", "''");
                    string releaseDate = update.ReleaseDate == default(DateTime) ? "NULL" : "\'" + update.ReleaseDate.ToString("yyyyMMdd") + "\'";
                    string kbNumber = (update.KbNumber ?? "N/A").Replace("'", "''");
                    string updateName = update.UpdateName != null ? update.UpdateName.Replace("'", "''") : "";
                    string url = update.Url != null ? update.Url.Replace("'", "''") : "";

                    sb.AppendLine($"INSERT INTO #SQLBuildsData ([MajorVersion], [BuildNumber], [KbNumber], [ReleaseDate], [UpdateName], [Url]) VALUES (N'{major}', N'{buildNumber}', N'{kbNumber}', {releaseDate}, N'{updateName}', N'{url}');");
                }
            }

            sb.AppendLine();
            sb.AppendLine("SELECT * FROM #SQLBuildsData;");
            sb.AppendLine();

            // Copy the generated script to the clipboard
            try
            {
                Clipboard.SetDataObject(sb.ToString());
                StatusText.Text = "TSQL script copied to clipboard.";
            }
            catch (Exception ex) { ShowActionError(ex); }
        }
    }
}
