namespace AxialSqlTools
{
    using Microsoft.SqlServer.Management.UI.VSIntegration;
    using Microsoft.VisualStudio.Shell;
    using System;
    using Microsoft.Data.SqlClient;
    using System.Diagnostics.CodeAnalysis;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows;
    using System.Windows.Media; 
    using System.Windows.Controls;
    using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;
    using System.Net.Http;
    using System.Collections.Generic;
    using System.Windows.Input;
    using System.Linq;
    using ScottPlot.WPF;
    using Plot = ScottPlot.Plot;
    using static HealthDashboardServerMetric;
    using System.Diagnostics;
    using System.IO;
    using System.Reflection;
    using System.Windows.Navigation;
    using static AxialSqlTools.AxialSqlToolsPackage;

    /// <summary>
    /// Interaction logic for HealthDashboard_ServerControl.
    /// </summary>
    public partial class HealthDashboard_ServerControl : UserControl
    {
        private readonly ToolWindowThemeController _themeController;
        private ServerHealthChartTheme _chartTheme;
        private readonly Dictionary<string, int> _waitColors = new Dictionary<string, int>(StringComparer.Ordinal);
        private ServerHealthWaitHistory waitsStatsAggregator = new ServerHealthWaitHistory();

        private WpfPlot[] ChartViews => new[] { DiskInfoModel, WaitStatsModel, PerfChart_CpuUtilization,
            PerfChart_UserConnections, PerfChart_BatchRequests, PerfChart_PageLifeExpectancy, PerfChart_PageReads,
            PerfChart_Transactions, PerfChart_LockWaits, PerfChart_MemoryGrantsPending, PerfChart_TotalServerMemory,
            BackupTimelineModel, BackupSizeModel, AgentJobsTimelineModel };

        public class PerformanceSample
        {
            public DateTime Timestamp { get; set; }
            public double CpuUtilization { get; set; }
            public double UserConnections { get; set; }
            public double BlockedRequests { get; set; }
            public double BatchRequestsSec { get; set; }
            public double SqlCompilationsSec { get; set; }
            public double PageLifeExpectancy { get; set; }
            public double PageReadsSec { get; set; }
            public double PageWritesSec { get; set; }
            public double LogFlushesSec { get; set; }
            public double TransactionsSec { get; set; }
            public double LockWaitsSec { get; set; }
            public double MemoryGrantsPending { get; set; }
            public double TotalServerMemoryMb { get; set; }
            public double TargetServerMemoryMb { get; set; }
        }

        public string connectionString = null;
        public DateTime lastRefresh = new DateTime(2000, 1, 1, 0, 0, 0);
        private CancellationTokenSource _cancellationTokenSource;
        private bool _disposed = false;
        private bool _monitoringStarted = false;
        private bool _versionCheckCompleted = false;
        private bool _newVersionAvailable = false;
        private string _newVersionURL;
        private readonly List<PerformanceSample> _performanceSamples = new List<PerformanceSample>();
        private const int PerformanceWindowMinutes = 15;


        private HealthDashboardServerMetric prev_metrics = new HealthDashboardServerMetric();

        private SettingsManager.HealthDashboardServerQueryTexts QueryLibrary = new SettingsManager.HealthDashboardServerQueryTexts();

        public ToolWindowPane userControlOwner;

        public void Dispose(bool disposing)
        {
            _cancellationTokenSource?.Cancel();

            _monitoringStarted = false;

            if (!_disposed)
            {
                if (disposing)
                {
                    _themeController.Dispose();
                    foreach (var chart in ChartViews) chart.Plot.Dispose();
                }

                // TODO: Free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: Set large fields to null
                _disposed = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="HealthDashboard_ServerControl"/> class.
        /// </summary>
        public HealthDashboard_ServerControl()
        {
            // ScottPlot's BAML requests SkiaSharp.Views.WPF without a version. In SSMS,
            // that bind can miss the extension's probing path. Load the packaged assembly
            // before WPF reads the template; do not resolve it relative to the host executable.
            Assembly.LoadFrom(Path.Combine(
                Path.GetDirectoryName(typeof(HealthDashboard_ServerControl).Assembly.Location),
                "SkiaSharp.Views.WPF.dll"));
            this.InitializeComponent();
            foreach (var chart in ChartViews)
            {
                chart.UserInputProcessor.Disable();
                chart.Menu = null;
                chart.MouseMove += Chart_MouseMove;
                chart.MouseLeave += Chart_MouseLeave;
                var tooltip = new ToolTip();
                tooltip.SetResourceReference(System.Windows.Controls.Control.BackgroundProperty, "AxialThemeBackgroundBrush");
                tooltip.SetResourceReference(System.Windows.Controls.Control.ForegroundProperty, "AxialThemeForegroundBrush");
                tooltip.SetResourceReference(System.Windows.Controls.Control.BorderBrushProperty, "AxialThemeBorderBrush");
                chart.ToolTip = tooltip;
            }
            _themeController = new ToolWindowThemeController(this, ApplyThemeBrushResources);
            UpdatePerformanceCharts();

            BackupTimelinePeriodNumberTextBox.Text = "1";
            AgentJobsTimelinePeriodNumberTextBox.Text = "1";

            AgentJobsUnsuccessfulOnly.IsChecked = true;

            DatabaseBackupHistoryIncludeFULL.IsChecked = true;
            DatabaseBackupHistoryIncludeDIFF.IsChecked = true;
            DatabaseBackupHistoryIncludeLOG.IsChecked = true;
           
        }

        private void ApplyThemeBrushResources()
        {
            ToolWindowThemeResources.ApplySharedTheme(this);
            var background = VsThemeBrushResolver.GetBrushColor(GetThemeBrush("AxialThemeBackgroundBrush", SystemColors.WindowBrush), System.Windows.Media.Colors.White);
            var foreground = VsThemeBrushResolver.GetBrushColor(GetThemeBrush("AxialThemeForegroundBrush", SystemColors.WindowTextBrush), System.Windows.Media.Colors.Black);
            // Start from the original palette on every theme change, avoiding progressive color drift.
            string[] palette = { "#0072B2", "#D55E00", "#009E73", "#CC79A7", "#8B78DB", "#A68A00", "#708090" };
            _chartTheme = new ServerHealthChartTheme
            {
                Background = ToPlotColor(background), Foreground = ToPlotColor(foreground),
                Border = ToPlotColor(VsThemeBrushResolver.GetBrushColor(GetThemeBrush("AxialThemeBorderBrush", SystemColors.ActiveBorderBrush), foreground)),
                Grid = ToPlotColor(VsThemeBrushResolver.GetBrushColor(GetThemeBrush("AxialThemeSubtleBorderBrush", SystemColors.InactiveBorderBrush), foreground)),
                HighContrast = SystemParameters.HighContrast,
                Series = palette.Select(hex => ToPlotColor(SystemParameters.HighContrast ? foreground
                    : VsThemeBrushResolver.EnsureTextContrast((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(hex), background, foreground))).ToArray()
            };
            foreach (var view in ChartViews)
            {
                ServerHealthCharts.ApplyTheme(view.Plot, _chartTheme);
                view.Refresh();
            }
        }

        private static ScottPlot.Color ToPlotColor(System.Windows.Media.Color color) => new ScottPlot.Color(color.R, color.G, color.B, color.A);

        private void ShowPlot(WpfPlot view, Plot plot)
        {
            ServerHealthCharts.ApplyTheme(plot, _chartTheme);
            view.Reset(plot); // Reset also disposes the previous plot.
            view.Refresh();
        }

        private void Chart_MouseMove(object sender, MouseEventArgs e)
        {
            var view = (WpfPlot)sender;
            var tooltip = (ToolTip)view.ToolTip;
            var point = view.Plot.GetCoordinates(view.GetPlotPixelPosition(e));
            string text = ServerHealthCharts.HoverText(view.Plot, point);
            tooltip.Content = text;
            if (string.IsNullOrEmpty(text)) tooltip.IsOpen = false;
        }

        private void Chart_MouseLeave(object sender, MouseEventArgs e)
        {
            ((ToolTip)((WpfPlot)sender).ToolTip).IsOpen = false;
        }

        public void StartMonitoring()
        {
            if (_monitoringStarted) return;

            waitsStatsAggregator = new ServerHealthWaitHistory();

            UpdateUI(0, new HealthDashboardServerMetric { }, true);

            var connectionInfo = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer(inMaster: true);

            if (connectionInfo == null) return;

            if (!string.IsNullOrEmpty(connectionInfo.FullConnectionString))
            {                
                connectionString = connectionInfo.FullConnectionString;

                _cancellationTokenSource = new CancellationTokenSource();
                _ = Task.Run(() => MonitorServerLoopAsync(_cancellationTokenSource.Token));
                _ = Task.Run(() => UpdateLastPullLoopAsync(_cancellationTokenSource.Token));

                _monitoringStarted = true;

            }

        }

        private async Task MonitorServerLoopAsync(CancellationToken cancellationToken)
        {
            try
            {
                int i = 0;

                while (!cancellationToken.IsCancellationRequested)
                {
                    var metrics = await MetricsService.FetchServerMetricsAsync(connectionString, prev_metrics);

                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
                    UpdateUI(i, metrics, false);
                    if (metrics.Completed && !metrics.HasException) prev_metrics = metrics;

                    await Task.Delay(3000, cancellationToken);
                    i += 1;
                }
            }
            catch (OperationCanceledException)
            {
            }
        }

        private async Task UpdateLastPullLoopAsync(CancellationToken cancellationToken)
        {
            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
                    UpdateTheLastPullDate();
                    await Task.Delay(1000, cancellationToken);
                }
            }
            catch (OperationCanceledException)
            {
            }
        }

        public void UpdateUI(int i, HealthDashboardServerMetric metrics, bool doEmpty)
        {
            if (doEmpty == false)
            {
                if (metrics.HasException)
                {
                    LabelInternalException.Content = metrics.ExecutionException;
                    LabelInternalException.SetResourceReference(System.Windows.Controls.Control.ForegroundProperty, "AxialThemeStatusErrorBrush");
                    LabelInternalException.FontWeight = System.Windows.FontWeights.Bold;

                    return;
                }
                else if (!metrics.Completed)
                {
                    return;
                }
            }

            bool ServerHasIssues = false;

            LabelInternalException.Content = "";

            if (metrics.spWhoIsActiveExists)
            {
                LinkRunSpWhoIsActive.Visibility = Visibility.Visible;
                LinkDeploySpWhoIsActive.Visibility = Visibility.Collapsed;
            } else
            {
                LinkRunSpWhoIsActive.Visibility = Visibility.Collapsed;
                LinkDeploySpWhoIsActive.Visibility = Visibility.Visible;
            }            

            Label_ServerName.Content = metrics.ServerName;
            Label_ServiceName.Content = metrics.ServiceName;
            Label_ServerVersion.Content = metrics.ServerVersion;

            //- Uptime
            TimeSpan difference = DateTime.Now - metrics.UtcStartTime.ToLocalTime();
            string formattedDifference = $"{difference.Days} days {difference.Hours} hours {difference.Minutes} minutes";

            Label_Uptime.Content = formattedDifference;

            Label_ResponseTime.Content = metrics.ServerResponseTimeMs.ToString();
            
            //-----------------------------------------
            // server memory
            string usedMemory = FormatBytesToGB(metrics.PerfCounter_TotalServerMemory * 1024);
            string targetMemory = FormatBytesToGB(metrics.PerfCounter_TargetServerMemory * 1024);
            string serverMemory = FormatBytesToGB(metrics.ServerMemoryTotal * 1024);

            Label_MemoryInfo.Content = $"{usedMemory} / {targetMemory} / {serverMemory}";
            //\\-----------------------------------------

            TimeSpan pleTimeSpan = TimeSpan.FromSeconds(metrics.PerfCounter_PLE);
            if (pleTimeSpan.TotalMinutes > 60)
                formattedDifference = $"{(int)pleTimeSpan.TotalHours} hours {pleTimeSpan.Minutes} minutes {pleTimeSpan.Seconds} sec.";
            else formattedDifference = $"{(int)pleTimeSpan.TotalMinutes} minutes {pleTimeSpan.Seconds} sec.";
            Label_PLE.Content = formattedDifference;

            Label_CpuLoad.Content = metrics.CpuUtilization + "%";
            Label_ConnectionCount.Content = metrics.ConnectionCountTotal.ToString();
            Label_ConnectionCountEnc.Content = metrics.ConnectionCountEnc.ToString();

            if (metrics.BlockedRequestsCount > 0)
            {
                Label_BlockedRequestCount.SetResourceReference(System.Windows.Controls.Control.ForegroundProperty, "AxialThemeStatusErrorBrush");
                Label_BlockedRequestCount.Content = metrics.BlockedRequestsCount.ToString();

                ServerHasIssues = true;

            }
            else {
                Label_BlockedRequestCount.SetResourceReference(System.Windows.Controls.Control.ForegroundProperty, "AxialThemeForegroundBrush");
                Label_BlockedRequestCount.Content = "-";
            }

            if (metrics.BlockingTotalWaitTime > 0)
            {
                Label_BlockedTotalWaitTime.SetResourceReference(System.Windows.Controls.Control.ForegroundProperty, "AxialThemeStatusErrorBrush");
                Label_BlockedTotalWaitTime.Content = metrics.BlockingTotalWaitTime.ToString();
            }
            else
            {
                Label_BlockedTotalWaitTime.SetResourceReference(System.Windows.Controls.Control.ForegroundProperty, "AxialThemeForegroundBrush");
                Label_BlockedTotalWaitTime.Content = "-";
            }

            //-------------------------------------------------
            Label_DatabaseStatus.SetResourceReference(System.Windows.Controls.Control.ForegroundProperty, "AxialThemeForegroundBrush");
            if (metrics.CountUserDatabasesTotal == 0)
                Label_DatabaseStatus.Content = "no user databases";
            else if (metrics.CountUserDatabasesTotal == metrics.CountUserDatabasesOkay)
            {
                Label_DatabaseStatus.Content = $"OK - {metrics.CountUserDatabasesTotal} database(s)";
            } else
            {
                Label_DatabaseStatus.SetResourceReference(System.Windows.Controls.Control.ForegroundProperty, "AxialThemeStatusErrorBrush");
                Label_DatabaseStatus.Content = $"{metrics.CountUserDatabasesOkay} out of {metrics.CountUserDatabasesTotal} available";
                ServerHasIssues = true;
            }

            if (metrics.Iteration > 2)
            {
                Label_BatchRequestsSec.Content = FormatRate(metrics.PerfCounter_BatchRequestsSec);
                Label_SQLCompilationsSec.Content = FormatRate(metrics.PerfCounter_SQLCompilationsSec);
            }

            //-------------------------------------------------
            if (metrics.AlwaysOn_Exists)
            {
                string agStatus = "HEALTHY";
                if (metrics.AlwaysOn_Health == 2)
                {
                    Label_AlwaysOnHealth.SetResourceReference(System.Windows.Controls.Control.ForegroundProperty, "AxialThemeStatusSuccessBrush");
                } else
                {
                    Label_AlwaysOnHealth.SetResourceReference(System.Windows.Controls.Control.ForegroundProperty, "AxialThemeStatusErrorBrush");
                    if (metrics.AlwaysOn_Health == 1)
                        agStatus = "PARTIALLY HEALTHY";
                    else agStatus = "NOT HEALTHY";

                    ServerHasIssues = true;
                }

                string agLatency = "";
                if (metrics.AlwaysOn_MaxLatency > 0)
                {
                    TimeSpan agTimeSpan = TimeSpan.FromMilliseconds(metrics.AlwaysOn_MaxLatency);
                    if (agTimeSpan.TotalMinutes > 1)
                        agLatency = $"{(int)agTimeSpan.Minutes} minutes {agTimeSpan.Seconds} sec.";
                    else 
                        agLatency = $"{(int)agTimeSpan.TotalSeconds} sec. {agTimeSpan.Milliseconds} ms.";                   
                }              

                Label_AlwaysOnHealth.Content = $"{agStatus} (latency: {agLatency})";
                Label_AlwaysOnLogSendQueue.Content = FormatBytesToMB(metrics.AlwaysOn_TotalLogSentQueueSize);
                Label_AlwaysOnRedoQueue.Content = FormatBytesToMB(metrics.AlwaysOn_TotalRedoQueueSize);
            }
            else {
                Label_AlwaysOnHealth.Content = "-";
                Label_AlwaysOnLogSendQueue.Content = "-";
                Label_AlwaysOnRedoQueue.Content = "-";
            }
            
           


            Label_DataFileSizeGb.Content = FormatBytesToGB(metrics.PerfCounter_DataFileSize * 1024);

            long usedLogFilePercent;
            if (metrics.PerfCounter_LogFileSize != 0)
                usedLogFilePercent = metrics.PerfCounter_UsedLogFileSize * 100 / metrics.PerfCounter_LogFileSize;
            else usedLogFilePercent = 0; 

            Label_LogFileSizeGb.Content = FormatBytesToGB(metrics.PerfCounter_LogFileSize * 1024) + " | " + usedLogFilePercent.ToString() + "% used";

            //--------------------------------------------------------------------
            // Wait Stats info graph

            if (!doEmpty)
                waitsStatsAggregator.Update(metrics.WaitStatsInfo.ToDictionary(x => x.WaitName, x => x.WaitSec, StringComparer.Ordinal),
                    DateTime.Now, metrics.UtcStartTime);
            ShowPlot(WaitStatsModel, ServerHealthCharts.Waits(waitsStatsAggregator.Snapshot(), DateTime.Now, _waitColors));
            UpdateDiskChart(metrics.DisksInfo);

            //--------------------------------------------------------------------
            //--------------------------------------------------------------------
            if (!doEmpty && metrics != null && metrics.Completed && !metrics.HasException && metrics.Iteration > 1)
            {
                AddPerformanceSample(metrics);
            }

            lastRefresh = DateTime.Now;

            //Let user know that there is an issue by blinking the title
            if (ServerHasIssues && CheckBox_StopBlinking.IsChecked == false)
            {
                if (userControlOwner.Caption.EndsWith("(!)"))
                    userControlOwner.Caption = metrics.ServerName + " - NOT OK";
                else
                    userControlOwner.Caption = "(!) " + metrics.ServerName + " - NOT OK (!)";
            } else if (ServerHasIssues)
                userControlOwner.Caption = metrics.ServerName + " - NOT OK";
            else
                if (string.IsNullOrEmpty(metrics.ServerName))
                    userControlOwner.Caption = "Health Dashboard | Server";
                else
                    userControlOwner.Caption = metrics.ServerName + " - OK";


            //-------------------------------------------------------------------
            if (!_versionCheckCompleted && !string.IsNullOrEmpty( metrics.ServerVersionShort))
            {

                try
                {
                    SQLBuildsData sqlBuilds = AxialSqlToolsPackage.PackageInstance.SQLBuildsDataInfo;

                    if (System.Version.TryParse(metrics.ServerVersionShort, out System.Version currentBuildVersion))
                    {

                        string targetSqlVersion = "";
                        switch (currentBuildVersion.Major)
                        {
                            case 17: targetSqlVersion = "2025"; break;
                            case 16: targetSqlVersion = "2022"; break;
                            case 15: targetSqlVersion = "2019"; break;
                            case 14: targetSqlVersion = "2017"; break;
                            case 13: targetSqlVersion = "2016"; break;
                            case 12: targetSqlVersion = "2014"; break;
                            case 11: targetSqlVersion = "2012"; break;
                        }

                        if (!string.IsNullOrEmpty(targetSqlVersion))
                        {
                            if (sqlBuilds.Builds.TryGetValue(targetSqlVersion, out List<SQLVersionInfo> builds))
                            {

                                SQLVersionInfo latestVersionInfo = builds.OrderByDescending(info => info.BuildNumber).FirstOrDefault();
                                if (latestVersionInfo != null)
                                {
                                    if (latestVersionInfo.BuildNumber > currentBuildVersion)
                                    {
                                        _newVersionAvailable = true;
                                        _newVersionURL = latestVersionInfo.Url;

                                        HyperlinkOpenNewVersionLink.Text = $"A new version is now available! {latestVersionInfo.UpdateName} released on: {latestVersionInfo.ReleaseDate.ToShortDateString()}";

                                    }                                    
                                }
                            }
                        }

                        if (_newVersionAvailable)
                        {
                            TextBlockNewVersion.Visibility = Visibility.Visible;
                        }
                        else
                        {
                            TextBlockNewVersion.Visibility = Visibility.Collapsed;
                        }
                    }                  
                    

                } catch {
                    TextBlockNewVersion.Visibility = Visibility.Collapsed;
                }                

                _versionCheckCompleted = true;
            }

        }

        private void UpdateDiskChart(List<DiskInfo> disks)
        {
            var plot = ServerHealthCharts.Create("Volume utilization");
            plot.XLabel("GB", 11);
            var used = new List<ScottPlot.Bar>();
            var free = new List<ScottPlot.Bar>();
            for (int i = 0; i < disks.Count; i++)
            {
                var disk = disks[i];
                used.Add(new ScottPlot.Bar { Position = i, Value = disk.UsedSpaceGb, Orientation = ScottPlot.Orientation.Horizontal,
                    Label = disk.UsedSpaceGb.ToString("N0") + " GB", CenterLabel = true });
                free.Add(new ScottPlot.Bar { Position = i, ValueBase = disk.UsedSpaceGb,
                    Value = disk.UsedSpaceGb + disk.FreeSpaceGb, Orientation = ScottPlot.Orientation.Horizontal,
                    Label = disk.FreeSpaceGb.ToString("N0") + " GB", CenterLabel = true });
            }
            ServerHealthCharts.Bars(plot, used, "Used", 1);
            ServerHealthCharts.Bars(plot, free, "Free", 2);
            ServerHealthCharts.LegendBelow(plot);
            plot.Axes.Left.SetTicks(Enumerable.Range(0, disks.Count).Select(x => (double)x).ToArray(), disks.Select(x => x.VolumeDescription).ToArray());
            plot.Axes.SetLimitsX(0, Math.Max(1, disks.Select(x => (double)x.TotalCapacityGb).DefaultIfEmpty(0).Max() * 1.05));
            plot.Axes.SetLimitsY(-0.5, Math.Max(0.5, disks.Count - 0.5));
            ServerHealthCharts.SetHover(plot, point =>
            {
                int index = (int)Math.Round(point.Y);
                return index >= 0 && index < disks.Count ? disks[index].VolumeDescription
                    + $"\nUsed: {disks[index].UsedSpaceGb:N0} GB\nFree: {disks[index].FreeSpaceGb:N0} GB" : null;
            });
            ShowPlot(DiskInfoModel, plot);
        }

        private static string FormatRate(double value) => double.IsNaN(value) ? "Collecting..." : value.ToString("N1");


        private void AddPerformanceSample(HealthDashboardServerMetric metrics)
        {
            var sample = new PerformanceSample
            {
                Timestamp = DateTime.Now,
                CpuUtilization = metrics.CpuUtilization,
                UserConnections = metrics.ConnectionCountTotal,
                BlockedRequests = metrics.BlockedRequestsCount,
                BatchRequestsSec = metrics.PerfCounter_BatchRequestsSec,
                SqlCompilationsSec = metrics.PerfCounter_SQLCompilationsSec,
                PageLifeExpectancy = metrics.PerfCounter_PLE,
                PageReadsSec = metrics.PerfCounter_PageReadsSec,
                PageWritesSec = metrics.PerfCounter_PageWritesSec,
                LogFlushesSec = metrics.PerfCounter_LogFlushesSec,
                TransactionsSec = metrics.PerfCounter_TransactionsSec,
                LockWaitsSec = metrics.PerfCounter_LockWaitsSec,
                MemoryGrantsPending = metrics.PerfCounter_MemoryGrantsPending,
                TotalServerMemoryMb = metrics.PerfCounter_TotalServerMemory / 1024.0,
                TargetServerMemoryMb = metrics.PerfCounter_TargetServerMemory / 1024.0
            };

            _performanceSamples.Add(sample);

            DateTime threshold = DateTime.Now.AddMinutes(-PerformanceWindowMinutes);
            _performanceSamples.RemoveAll(s => s.Timestamp < threshold);

            UpdatePerformanceCharts();
        }

        private void UpdateTimeSeries(WpfPlot view, string title, string units, string[] labels,
            params Func<PerformanceSample, double>[] selectors)
        {
            var samples = _performanceSamples.OrderBy(x => x.Timestamp).ToArray();
            var xs = samples.Select(x => x.Timestamp.ToOADate()).ToArray();
            var values = selectors.Select(select => samples.Select(select).ToArray()).ToArray();
            ShowPlot(view, ServerHealthCharts.TimeSeries(title, units, labels, xs, values, DateTime.Now, view == PerfChart_CpuUtilization));
        }

        private void UpdatePerformanceCharts()
        {
            UpdateTimeSeries(PerfChart_CpuUtilization, "SQL Server CPU", "%", new[] { "CPU" }, s => s.CpuUtilization);
            UpdateTimeSeries(PerfChart_BatchRequests, "SQL activity", "/sec", new[] { "Batches", "Compilations" }, s => s.BatchRequestsSec, s => s.SqlCompilationsSec);
            UpdateTimeSeries(PerfChart_UserConnections, "Connections & blocking", "count", new[] { "Connections", "Blocked requests" }, s => s.UserConnections, s => s.BlockedRequests);
            UpdateTimeSeries(PerfChart_TotalServerMemory, "SQL Server memory", "MB", new[] { "Total", "Target" }, s => s.TotalServerMemoryMb, s => s.TargetServerMemoryMb);
            UpdateTimeSeries(PerfChart_PageLifeExpectancy, "Page life expectancy", "seconds", new[] { "PLE" }, s => s.PageLifeExpectancy);
            UpdateTimeSeries(PerfChart_PageReads, "Page I/O", "pages/sec", new[] { "Reads", "Writes" }, s => s.PageReadsSec, s => s.PageWritesSec);
            UpdateTimeSeries(PerfChart_Transactions, "Transactions & log flushes", "/sec", new[] { "Transactions", "Log flushes" }, s => s.TransactionsSec, s => s.LogFlushesSec);
            UpdateTimeSeries(PerfChart_LockWaits, "Lock waits", "waits/sec", new[] { "Lock waits" }, s => s.LockWaitsSec);
            UpdateTimeSeries(PerfChart_MemoryGrantsPending, "Memory grants pending", "count", new[] { "Pending" }, s => s.MemoryGrantsPending);
        }

        public static string FormatBytesToMB(long bytes)
        {
            // Convert bytes to gigabytes (GB)
            double megabytes = bytes / 1024.0 / 1024.0; 

            // Format the gigabytes to a string with one decimal place
            return $"{megabytes:#,0.0} MB";
        }

        public static string FormatBytesToGB(long bytes)
        {
            // Convert bytes to gigabytes (GB)
            double gigabytes = bytes / 1073741824.0; // 1024^3

            // Format the gigabytes to a string with one decimal place
            return $"{gigabytes:#,0.0} GB";
        }

        public void UpdateTheLastPullDate()
        {
            TimeSpan elapsedTime = DateTime.Now - lastRefresh;

            // Get the total number of seconds (including fractions of a second)
            int totalSeconds = (int)elapsedTime.TotalSeconds;

            LastUpdateLabel.Content = "Last updated: " + lastRefresh.ToString() + " | " + totalSeconds + " sec. ago.";

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
                "HealthDashboard_Server");

        }

        void OpenNewQueryWindowAndExecute(string QueryText, bool Execute = true)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var connectionInfo = ScriptFactoryAccess.GetCurrentConnectionInfo();
            ServiceCache.ScriptFactory.CreateNewBlankScript(ScriptType.Sql, connectionInfo.ActiveConnectionInfo, null);

            EnvDTE.TextDocument doc = (EnvDTE.TextDocument)ServiceCache.ExtensibilityModel.Application.ActiveDocument.Object(null);

            doc.EndPoint.CreateEditPoint().Insert(QueryText);

            if (Execute)
                ServiceCache.ExtensibilityModel.Application.ActiveDocument.DTE.ExecuteCommand("Query.Execute");
        }

        private void buttonBlockedRequests_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread(); // helps to suppress build warning
            OpenNewQueryWindowAndExecute(QueryLibrary.BlockingRequests);
        }

        private void buttonDatabaseLogInfo_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            OpenNewQueryWindowAndExecute(QueryLibrary.DatabaseLogUsageInfo);
        }

        private void buttonUserDatabasesInfo_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            OpenNewQueryWindowAndExecute(QueryLibrary.DatabaseInfo);
        }

        private void buttonAlwaysOn_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            OpenNewQueryWindowAndExecute(QueryLibrary.AlwaysOnStatus);
        }

        private void buttonRunSpWhoIsActive_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            OpenNewQueryWindowAndExecute(QueryLibrary.spWhoIsActive);
        }

        private async void buttonDeploySpWhoIsActive_Click(object sender, RoutedEventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            string FullCode = "USE [master]\nGO\n";
            bool error = false;

            // sp_WhoIsActive source
            string url = "https://raw.githubusercontent.com/amachanic/sp_whoisactive/master/sp_WhoIsActive.sql";
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    string content = await client.GetStringAsync(url);

                    FullCode += content;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "An error occurred");
                error = true;
            }

            if (!error)
                OpenNewQueryWindowAndExecute(FullCode, Execute: false);
        }

        private void buttonDetailedBackupInfo_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            OpenNewQueryWindowAndExecute(QueryLibrary.DatabaseBackupDetailedInfo);
        }



        private void BackupTimelineModelRefresh_Click(object sender, RoutedEventArgs e)
        {

            //var ci = ScriptFactoryAccess.GetCurrentConnectionInfo();

            int BackupHistoryPeriod = 1;
            int.TryParse(BackupTimelinePeriodNumberTextBox.Text, out BackupHistoryPeriod);

            string sourceQuery = @"
            SELECT 
                database_name, 
                backup_start_date, 
                backup_finish_date, 
                CASE type
                    WHEN 'D' THEN 0.0
                    WHEN 'I' THEN 0.1
                    WHEN 'L' THEN 0.2
                END AS BackupType
            FROM msdb.dbo.backupset
            WHERE 
                backup_start_date > GETDATE() - @BackupHistoryPeriod
                AND (@Include_Full = 1 OR type <> 'D')
                AND (@Include_Diff = 1 OR type <> 'I')
                AND (@Include_Log = 1 OR type <> 'L')
            ORDER BY
                database_name DESC, backup_start_date;";

            var MyModel = ServerHealthCharts.Create("Database Backups: Frequency and Durations Analysis");
            var timelineDetails = new List<Tuple<double, double, double, string>>();

            var customLabels = new Dictionary<double, string>();
            var databaseIndex = new Dictionary<string, double>();

            using (SqlConnection sourceConn = new SqlConnection(connectionString))
            {
                sourceConn.Open();
                using (SqlCommand cmd = new SqlCommand(sourceQuery, sourceConn))
                {
                    cmd.Parameters.AddWithValue("BackupHistoryPeriod", BackupHistoryPeriod);
                    cmd.Parameters.AddWithValue("Include_Full", DatabaseBackupHistoryIncludeFULL.IsChecked.GetValueOrDefault());
                    cmd.Parameters.AddWithValue("Include_Diff", DatabaseBackupHistoryIncludeDIFF.IsChecked.GetValueOrDefault());
                    cmd.Parameters.AddWithValue("Include_Log", DatabaseBackupHistoryIncludeLOG.IsChecked.GetValueOrDefault());
                    
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {

                            var databaseName = reader.GetString(0);                            

                            double dbIndex = -1;
                            if (!databaseIndex.TryGetValue(databaseName, out dbIndex))
                            {
                                dbIndex = databaseIndex.Count;
                                databaseIndex.Add(databaseName, dbIndex);
                            }

                            if (!customLabels.ContainsKey(dbIndex))
                                customLabels.Add(dbIndex, databaseName);


                            var startDate = reader.GetDateTime(1).ToOADate();
                            var finishDate = reader.GetDateTime(2).ToOADate();

                            var backupType = (double)reader.GetDecimal(3);

                            int colorIndex = backupType == 0.1 ? 2 : backupType == 0.2 ? 1 : 0;
                            string kind = backupType == 0.1 ? "DIFF" : backupType == 0.2 ? "LOG" : "FULL";
                            dbIndex += backupType;
                            var line = ServerHealthCharts.Line(MyModel, new[] { startDate, finishDate }, new[] { dbIndex, dbIndex }, "", colorIndex);
                            line.MarkerSize = 4;
                            timelineDetails.Add(Tuple.Create(startDate, finishDate, dbIndex, databaseName + " · " + kind));

                        }
                    }
                }
            }

            ServerHealthCharts.FinishTimeline(MyModel, customLabels);
            SetTimelineHover(MyModel, timelineDetails);
            ShowPlot(BackupTimelineModel, MyModel);


            //---------------------------------------------------
            // Database Backup Sizes 

            string sourceQuerySizes = @"
            SELECT 
                database_name, 
                SUM(CAST(compressed_backup_size / 1024 / 1024 / 1024. AS NUMERIC (15, 2))) AS TotalBackupSize
            FROM msdb.dbo.backupset
            WHERE backup_start_date > GETDATE() - @BackupHistoryPeriod
                AND (@Include_Full = 1 OR type <> 'D')
                AND (@Include_Diff = 1 OR type <> 'I')
                AND (@Include_Log = 1 OR type <> 'L')
            GROUP BY database_name
            ORDER BY 2 DESC";

            var backupSizes = new List<KeyValuePair<string, double>>();

            using (SqlConnection sourceConn = new SqlConnection(connectionString))
            {
                sourceConn.Open();
                using (SqlCommand cmd = new SqlCommand(sourceQuerySizes, sourceConn))
                {
                    cmd.Parameters.AddWithValue("BackupHistoryPeriod", BackupHistoryPeriod);
                    cmd.Parameters.AddWithValue("Include_Full", DatabaseBackupHistoryIncludeFULL.IsChecked.GetValueOrDefault());
                    cmd.Parameters.AddWithValue("Include_Diff", DatabaseBackupHistoryIncludeDIFF.IsChecked.GetValueOrDefault());
                    cmd.Parameters.AddWithValue("Include_Log", DatabaseBackupHistoryIncludeLOG.IsChecked.GetValueOrDefault());
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {

                            var databaseName = reader.GetString(0);
                            var backupSize = (double)reader.GetDecimal(1);

                            backupSizes.Add(new KeyValuePair<string, double>(databaseName, backupSize));

                        }
                    }
                }
            }


            ShowPlot(BackupSizeModel, ServerHealthCharts.BackupSizes(backupSizes));

        }

        private void AgentJobsTimelineModelRefresh_Click(object sender, RoutedEventArgs e)
        {

            int AgentJobHistoryPeriod = 1;
            //TODO
            int.TryParse(AgentJobsTimelinePeriodNumberTextBox.Text, out AgentJobHistoryPeriod);

            string sourceQuery = @" 
            SELECT 
                j.name AS JobName,
                CONVERT(datetime, 
                    STUFF(STUFF(CAST(h.run_date AS varchar(8)), 5, 0, '-'), 8, 0, '-') + 
                    ' ' + 
                    STUFF(STUFF(RIGHT('000000' + CAST(h.run_time AS varchar(6)), 6), 3, 0, ':'), 6, 0, ':')
                ) AS StartTime,
                DATEADD(SECOND, h.run_duration / 10000 * 3600 + (h.run_duration / 100 % 100) * 60 + h.run_duration % 100, 
                    CONVERT(datetime, 
                        STUFF(STUFF(CAST(h.run_date AS varchar(8)), 5, 0, '-'), 8, 0, '-') + 
                        ' ' + 
                        STUFF(STUFF(RIGHT('000000' + CAST(h.run_time AS varchar(6)), 6), 3, 0, ':'), 6, 0, ':')
                    )
                ) AS EndTime,
                h.run_status
                --CASE h.run_status
                --    WHEN 3 THEN 'StoppedManually'
                --    WHEN 1 THEN 'Success'
                --    WHEN 0 THEN 'Failure'
                --    ELSE 'Unknown'
                --END AS ExecutionStatus
            FROM  msdb.dbo.sysjobs j
                JOIN msdb.dbo.sysjobhistory h ON j.job_id = h.job_id and step_id = 0
            WHERE 
                h.run_date >= CONVERT(int, CONVERT(varchar(8), GETDATE() - @AgentJobsHistoryPeriod, 112))
                AND h.run_date < CONVERT(int, CONVERT(varchar(8), GETDATE(), 112))
                AND (@UnsuccessfulOnly = 0 
                        OR h.run_status <> 1)
            ORDER BY 
                StartTime DESC;

            /* CREATE INDEX IDX_sysjobhistory_1
                ON sysjobhistory(job_id, step_id, run_date, run_time) 
	            WITH (DATA_COMPRESSION = PAGE, ONLINE = ON, MAXDOP = 4); */
            ";

            var MyModel = ServerHealthCharts.Create("Agent Jobs: Frequency and Durations Analysis");
            var timelineDetails = new List<Tuple<double, double, double, string>>();

            var customLabels = new Dictionary<double, string>();
            var jobIndex = new Dictionary<string, double>();

            using (SqlConnection sourceConn = new SqlConnection(connectionString))
            {
                sourceConn.Open();
                using (SqlCommand cmd = new SqlCommand(sourceQuery, sourceConn))
                {
                    cmd.Parameters.AddWithValue("AgentJobsHistoryPeriod", AgentJobHistoryPeriod);
                    cmd.Parameters.AddWithValue("UnsuccessfulOnly", AgentJobsUnsuccessfulOnly.IsChecked.GetValueOrDefault());
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {

                            var jobName = reader.GetString(0);

                            double jIndex = -1;
                            if (!jobIndex.TryGetValue(jobName, out jIndex))
                            {
                                jIndex = jobIndex.Count;
                                jobIndex.Add(jobName, jIndex);
                            }

                            if (!customLabels.ContainsKey(jIndex))
                                customLabels.Add(jIndex, jobName);


                            var startDate = reader.GetDateTime(1).ToOADate();
                            var finishDate = reader.GetDateTime(2).ToOADate();

                            var resultType = reader.GetInt32(3);

                            int colorIndex = resultType == 0 ? 1 : resultType == 1 ? 2 : resultType == 3 ? 3 : 6;
                            var line = ServerHealthCharts.Line(MyModel, new[] { startDate, finishDate }, new[] { jIndex, jIndex }, "", colorIndex);
                            line.MarkerSize = 4;
                            string status = resultType == 0 ? "Failed" : resultType == 1 ? "Succeeded" : resultType == 2 ? "Retry" : resultType == 3 ? "Canceled" : "In progress";
                            timelineDetails.Add(Tuple.Create(startDate, finishDate, jIndex, jobName + " · " + status));

                        }
                    }
                }
            }

            ServerHealthCharts.FinishTimeline(MyModel, customLabels);
            SetTimelineHover(MyModel, timelineDetails);
            ShowPlot(AgentJobsTimelineModel, MyModel);


        }

        private void BackupTimelinePeriodNumberTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Allow only numbers
            e.Handled = !int.TryParse(e.Text, out _);
        }

        private void BackupTimelinePeriodNumberTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Ensure the value stays within bounds
            if (int.TryParse(BackupTimelinePeriodNumberTextBox.Text, out int value))
            {
                if (value < 1) BackupTimelinePeriodNumberTextBox.Text = "1";
                else if (value > 14) BackupTimelinePeriodNumberTextBox.Text = "14";
            }
        }

        private void AgentJobsTimelinePeriodNumberTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Allow only numbers
            e.Handled = !int.TryParse(e.Text, out _);
        }

        private void AgentJobsTimelinePeriodNumberTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Ensure the value stays within bounds
            if (int.TryParse(AgentJobsTimelinePeriodNumberTextBox.Text, out int value))
            {
                if (value < 1) AgentJobsTimelinePeriodNumberTextBox.Text = "1";
                else if (value > 14) AgentJobsTimelinePeriodNumberTextBox.Text = "14";
            }
        }


        private void HyperlinkOpenNewVersionLink_Click(object sender, RoutedEventArgs e)
        {
            ToolWindowNavigation.OpenExternalUrl(_newVersionURL);
        }

        private void WikiLink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            ToolWindowNavigation.HandleRequestNavigate(e);
        }

        private static void SetTimelineHover(Plot plot, List<Tuple<double, double, double, string>> intervals)
        {
            ServerHealthCharts.SetHover(plot, point =>
            {
                var closest = intervals.OrderBy(x => Math.Abs(x.Item3 - point.Y)).ThenBy(x => Math.Abs(x.Item1 - point.X)).FirstOrDefault();
                if (closest == null) return "No history in this period";
                return closest.Item4 + "\nStart: " + DateTime.FromOADate(closest.Item1).ToString("g")
                    + "\nFinish: " + DateTime.FromOADate(closest.Item2).ToString("g")
                    + "\nDuration: " + TimeSpan.FromDays(closest.Item2 - closest.Item1).ToString();
            });
        }

        private Brush GetThemeBrush(string key, Brush fallback)
        {
            return TryFindResource(key) as Brush ?? fallback;
        }

    }
}
