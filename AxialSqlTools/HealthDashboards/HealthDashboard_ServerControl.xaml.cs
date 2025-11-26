namespace AxialSqlTools
{
    using Microsoft.SqlServer.Management.Smo.RegSvrEnum;
    using Microsoft.SqlServer.Management.UI.VSIntegration;
    using Microsoft.VisualStudio.Shell;
    using System;
    using System.Data.SqlClient;
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
    using static HealthDashboardServerMetric;
    using System.Diagnostics;
    using static AxialSqlTools.AxialSqlToolsPackage;
    using DocumentFormat.OpenXml.Bibliography;
    using DocumentFormat.OpenXml.Spreadsheet;
    using static AxialSqlTools.HealthDashboard_ServerControl;
    using ScottPlot;
    using ScottPlot.Plottables;
    using ScottPlot.WPF;
    using ScottPlot.Palettes;
    using Color = System.Drawing.Color;

    /// <summary>
    /// Interaction logic for HealthDashboard_ServerControl.
    /// </summary>
    public partial class HealthDashboard_ServerControl : UserControl
    {

        public class WaitsStatsAggregator
        {
            public List<WaitsInfo> previousWaitStats = new List<WaitsInfo>();
            private Dictionary<DateTime, List<WaitsInfo>> aggregatedWaitStats = new Dictionary<DateTime, List<WaitsInfo>>();
            //private readonly Timer timer;

            public WaitsStatsAggregator()
            {
                //// Set up a timer to reset the aggregation every minute
                //timer = new Timer(60000); // 60 seconds interval
                //timer.Elapsed += Timer_Elapsed;
                //timer.Start();
            }

            //private void Timer_Elapsed(object sender, ElapsedEventArgs e)
            //{
            //    // At the start of a new minute, reset the previousWaitStats
            //    previousWaitStats.Clear();
            //}

            public void UpdateWaitStats(List<WaitsInfo> currentWaitStats)
            {
                var timestamp = DateTime.Now;

                if (previousWaitStats.Count == 0)
                {
                    previousWaitStats = currentWaitStats;
                    return;
                }

                // Calculate the difference from the previous result
                var diff = currentWaitStats.Select(current => new WaitsInfo
                {
                    WaitName = current.WaitName,
                    WaitSec = previousWaitStats.Any(p => p.WaitName == current.WaitName) ? current.WaitSec - previousWaitStats.FirstOrDefault(p => p.WaitName == current.WaitName)?.WaitSec ?? 0 : current.WaitSec
                }).ToList();

                // Update the previousWaitStats for the next comparison
                previousWaitStats = currentWaitStats;

                // Grouping the values into minutes
                var minuteKey = new DateTime(timestamp.Year, timestamp.Month, timestamp.Day, timestamp.Hour, timestamp.Minute, 0);
                if (!aggregatedWaitStats.ContainsKey(minuteKey))
                {
                    aggregatedWaitStats[minuteKey] = new List<WaitsInfo>();
                }

                foreach (var item in diff)
                {
                    if (aggregatedWaitStats[minuteKey].Any(x => x.WaitName == item.WaitName))
                    {
                        aggregatedWaitStats[minuteKey].First(x => x.WaitName == item.WaitName).WaitSec += item.WaitSec;
                    }
                    else
                    {
                        aggregatedWaitStats[minuteKey].Add(item);
                    }
                }

                // Clean up entries older than 15 minutes
                var threshold = DateTime.Now.AddMinutes(-15);
                var keysToRemove = aggregatedWaitStats.Keys.Where(k => k < threshold).ToList();
                foreach (var key in keysToRemove)
                {
                    aggregatedWaitStats.Remove(key);
                }


            }

            // Method to retrieve the aggregated data (you can call this method to get the data for visualization)
            public Dictionary<DateTime, List<WaitsInfo>> GetAggregatedData()
            {
                return aggregatedWaitStats;
            }
        }

        public class PerformanceSample
        {
            public DateTime Timestamp { get; set; }
            public double CpuUtilization { get; set; }
            public double UserConnections { get; set; }
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

        private WaitsStatsAggregator waitsStatsAggregator = new WaitsStatsAggregator();

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
                    //...
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
            this.InitializeComponent();

            BackupTimelinePeriodNumberTextBox.Text = "1";
            AgentJobsTimelinePeriodNumberTextBox.Text = "1";

            AgentJobsUnsuccessfulOnly.IsChecked = true;

            DatabaseBackupHistoryIncludeFULL.IsChecked = true;
            DatabaseBackupHistoryIncludeDIFF.IsChecked = true;
            DatabaseBackupHistoryIncludeLOG.IsChecked = true;
           
        }

        public void StartMonitoring()
        {
            if (_monitoringStarted) return;

            waitsStatsAggregator = new WaitsStatsAggregator();

            UpdateUI(0, new HealthDashboardServerMetric { }, true);

            var connectionInfo = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer(inMaster: true);

            if (!string.IsNullOrEmpty(connectionInfo.FullConnectionString))
            {                
                connectionString = connectionInfo.FullConnectionString;

                _cancellationTokenSource = new CancellationTokenSource();

                ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
                {

                    int i = 0;

                    while (!_cancellationTokenSource.IsCancellationRequested)
                    {

                        var metrics = await MetricsService.FetchServerMetricsAsync(connectionString, prev_metrics);

                        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(_cancellationTokenSource.Token);

                        UpdateUI(i, metrics, false);

                        prev_metrics = metrics;

                        await System.Threading.Tasks.Task.Delay(3000, _cancellationTokenSource.Token); // Wait for 30 seconds before refreshing again

                        i += 1;
                    }

                });

                // update counter on the form with the last update time 
                ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
                {
                    while (!_cancellationTokenSource.IsCancellationRequested)
                    {
                        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(_cancellationTokenSource.Token);
                        UpdateTheLastPullDate();
                        await System.Threading.Tasks.Task.Delay(1000, _cancellationTokenSource.Token);
                    }
                });

                _monitoringStarted = true;

            }

        }

        public void UpdateUI(int i, HealthDashboardServerMetric metrics, bool doEmpty)
        {
            if (doEmpty == false)
            {
                if (metrics.HasException)
                {
                    LabelInternalException.Content = metrics.ExecutionException;
                    LabelInternalException.Foreground = Brushes.Red;
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
                Label_BlockedRequestCount.Foreground = Brushes.Red;
                Label_BlockedRequestCount.Content = metrics.BlockedRequestsCount.ToString();

                ServerHasIssues = true;

            }
            else {
                Label_BlockedRequestCount.Foreground = Brushes.Black;
                Label_BlockedRequestCount.Content = "-";
            }

            if (metrics.BlockingTotalWaitTime > 0)
            {
                Label_BlockedTotalWaitTime.Foreground = Brushes.Red;
                Label_BlockedTotalWaitTime.Content = metrics.BlockingTotalWaitTime.ToString();
            }
            else
            {
                Label_BlockedTotalWaitTime.Foreground = Brushes.Black;
                Label_BlockedTotalWaitTime.Content = "-";
            }

            //-------------------------------------------------
            Label_DatabaseStatus.Foreground = Brushes.Black;
            if (metrics.CountUserDatabasesTotal == 0)
                Label_DatabaseStatus.Content = "no user databases";
            else if (metrics.CountUserDatabasesTotal == metrics.CountUserDatabasesOkay)
            {
                Label_DatabaseStatus.Content = $"OK - {metrics.CountUserDatabasesTotal} database(s)";
            } else
            {
                Label_DatabaseStatus.Foreground = Brushes.Red;
                Label_DatabaseStatus.Content = $"{metrics.CountUserDatabasesOkay} out of {metrics.CountUserDatabasesTotal} available";
                ServerHasIssues = true;
            }

            if (metrics.Iteration > 2)
            {
                Label_BatchRequestsSec.Content = metrics.PerfCounter_BatchRequestsSec;
                Label_SQLCompilationsSec.Content = metrics.PerfCounter_SQLCompilationsSec;
            }

            //-------------------------------------------------
            if (metrics.AlwaysOn_Exists)
            {
                string agStatus = "HEALTHY";
                if (metrics.AlwaysOn_Health == 2)
                {
                    Label_AlwaysOnHealth.Foreground = Brushes.Black;
                } else
                {
                    Label_AlwaysOnHealth.Foreground = Brushes.Red;
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

            waitsStatsAggregator.UpdateWaitStats(metrics.WaitStatsInfo);
            UpdateWaitStatsGraph(waitsStatsAggregator.previousWaitStats, waitsStatsAggregator.GetAggregatedData());

            ////--------------------------------------------------------------------
            //// Disk info graph
            //var diskPlot = DiskInfoModel.Plot;
            //diskPlot.Clear();
            //diskPlot.Title("Volume(s) Utilization");
            //
            //double[] positions = Enumerable.Range(0, metrics.DisksInfo.Count).Select(ii => (double)ii).ToArray();
            //long[] usedValues = metrics.DisksInfo.Select(disk => disk.UsedSpaceGb).ToArray();
            //long[] freeValues = metrics.DisksInfo.Select(disk => disk.FreeSpaceGb).ToArray();
            //string[] labels = metrics.DisksInfo.Select(disk => disk.VolumeDescription).ToArray();
            //
            //var usedBars = diskPlot.AddBar(usedValues, positions);
            //usedBars.Horizontal = true;
            //usedBars.Label = "Used";
            //usedBars.FillColor = Color.LightPink;
            //usedBars.BorderColor = Color.Black;
            //usedBars.ValueFormatter = x => $"{x:0} Gb";
            //
            //var freeBars = diskPlot.AddBar(freeValues, positions);
            //freeBars.Horizontal = true;
            //freeBars.Label = "Free";
            //freeBars.FillColor = Color.LightBlue;
            //freeBars.BorderColor = Color.Black;
            //freeBars.ValueFormatter = x => $"{x:0} Gb";
            //
            //double offset = usedBars.BarWidth / 2;
            //usedBars.PositionOffset = -offset;
            //freeBars.PositionOffset = offset;
            //
            //diskPlot.Axes.Left.SetTicks(positions, labels);
            //diskPlot.XLabel("Gb");
            //diskPlot.SetAxisLimits(xMin: 0);
            //diskPlot.ShowLegend(Alignment.UpperRight);
            //
            //DiskInfoModel.Refresh();


            //--------------------------------------------------------------------
            //--------------------------------------------------------------------
            if (!doEmpty && metrics != null && metrics.Completed && !metrics.HasException)
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

        private void UpdateWaitStatsGraph(List<WaitsInfo> previousWaitStats, Dictionary<DateTime, List<WaitsInfo>> aggrData)
        {

            //var sortedKeys = aggrData.Keys.OrderBy(k => k).ToList();
            //var waitPlot = WaitStatsModel.Plot;
            //waitPlot.Clear();
            //waitPlot.Title("Real-time Wait Stats");
            //
            //double[] positions = Enumerable.Range(0, sortedKeys.Count).Select(i => (double)i).ToArray();
            //waitPlot.Axes.Left.SetTicks(positions, sortedKeys.Select(k => k.ToString("HH:mm")).ToArray());
            //waitPlot.SetAxisLimits(xMin: 0);
            //waitPlot.ShowLegend();
            //
            //var palette = new Category10();
            //double offsetStep = 0.1;
            //int seriesIndex = 0;
            //
            //foreach (var previousWaitStat in previousWaitStats)
            //{
            //    double[] values = sortedKeys
            //        .Select(key => aggrData[key].FirstOrDefault(ws => ws.WaitName == previousWaitStat.WaitName)?.WaitSec ?? 0)
            //        .Select(v => (double)v)
            //        .ToArray();
            //
            //    var barSeries = waitPlot.AddBar(values, positions);
            //    barSeries.Horizontal = true;
            //    barSeries.Label = previousWaitStat.WaitName;
            //    barSeries.FillColor = palette.GetColor(seriesIndex % palette.Count);
            //    barSeries.BorderColor = Color.Black;
            //    barSeries.PositionOffset = (seriesIndex - (previousWaitStats.Count / 2.0)) * offsetStep;
            //
            //    seriesIndex++;
            //}
            //
            //WaitStatsModel.Refresh();
        }

        private void AddPerformanceSample(HealthDashboardServerMetric metrics)
        {
            var sample = new PerformanceSample
            {
                Timestamp = DateTime.Now,
                CpuUtilization = metrics.CpuUtilization,
                UserConnections = metrics.ConnectionCountTotal,
                BatchRequestsSec = metrics.PerfCounter_BatchRequestsSec,
                SqlCompilationsSec = metrics.PerfCounter_SQLCompilationsSec,
                PageLifeExpectancy = metrics.PerfCounter_PLE,
                PageReadsSec = metrics.PerfCounter_PageReadsSec,
                PageWritesSec = metrics.PerfCounter_PageWritesSec,
                LogFlushesSec = metrics.PerfCounter_LogFlushesSec,
                TransactionsSec = metrics.PerfCounter_TransactionsSec,
                LockWaitsSec = metrics.PerfCounter_LockWaitsSec,
                MemoryGrantsPending = metrics.PerfCounter_MemoryGrantsPending,
                TotalServerMemoryMb = metrics.PerfCounter_TotalServerMemory / 1024.0
            };

            _performanceSamples.Add(sample);

            DateTime threshold = DateTime.Now.AddMinutes(-PerformanceWindowMinutes);
            _performanceSamples.RemoveAll(s => s.Timestamp < threshold);

            UpdatePerformanceCharts();
        }

        private void UpdateTimeSeriesPlot(WpfPlot targetPlot, string title, string yAxisTitle, Func<PerformanceSample, double> valueSelector, bool clampToZero = true)
        {
            //var orderedSamples = _performanceSamples.OrderBy(s => s.Timestamp).ToList();
            //
            //double[] xs = orderedSamples.Select(s => s.Timestamp.ToOADate()).ToArray();
            //double[] ys = orderedSamples.Select(valueSelector).ToArray();
            //
            //var plt = targetPlot.Plot;
            //plt.Clear();
            //plt.Title(title);
            //plt.XAxis.DateTimeFormat(true);
            //plt.YLabel(yAxisTitle);
            //plt.AddScatter(xs, ys, markerSize: 2, lineWidth: 2);
            //
            //if (clampToZero)
            //    plt.SetAxisLimits(yMin: 0);
            //
            //targetPlot.Refresh();
        }

        private void UpdatePerformanceCharts()
        {
            UpdateTimeSeriesPlot(PerfChart_CpuUtilization, "CPU Utilization (%)", "%", s => s.CpuUtilization);
            UpdateTimeSeriesPlot(PerfChart_UserConnections, "User Connections", "connections", s => s.UserConnections);
            UpdateTimeSeriesPlot(PerfChart_BatchRequests, "Batch Requests/sec", "requests/sec", s => s.BatchRequestsSec);
            UpdateTimeSeriesPlot(PerfChart_SqlCompilations, "SQL Compilations/sec", "compilations/sec", s => s.SqlCompilationsSec);
            UpdateTimeSeriesPlot(PerfChart_PageLifeExpectancy, "Page Life Expectancy", "seconds", s => s.PageLifeExpectancy, clampToZero: false);
            UpdateTimeSeriesPlot(PerfChart_PageReads, "Page Reads/sec", "pages/sec", s => s.PageReadsSec);
            UpdateTimeSeriesPlot(PerfChart_PageWrites, "Page Writes/sec", "pages/sec", s => s.PageWritesSec);
            UpdateTimeSeriesPlot(PerfChart_LogFlushes, "Log Flushes/sec", "flushes/sec", s => s.LogFlushesSec);
            UpdateTimeSeriesPlot(PerfChart_Transactions, "Transactions/sec", "transactions/sec", s => s.TransactionsSec);
            UpdateTimeSeriesPlot(PerfChart_LockWaits, "Lock Waits/sec", "waits/sec", s => s.LockWaitsSec);
            UpdateTimeSeriesPlot(PerfChart_MemoryGrantsPending, "Memory Grants Pending", "grants", s => s.MemoryGrantsPending);
            UpdateTimeSeriesPlot(PerfChart_TotalServerMemory, "Total Server Memory", "MB", s => s.TotalServerMemoryMb);
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
            /*
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

            var backupTimelinePlot = BackupTimelineModel.Plot;
            backupTimelinePlot.Clear();
            backupTimelinePlot.Title("Database Backups: Frequency and Durations Analysis");
            backupTimelinePlot.XAxis.DateTimeFormat(true);
            backupTimelinePlot.XLabel("Date");

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

                            var LineColor = Color.Blue;
                            if (backupType == 0.1)
                                LineColor = Color.Green;
                            if (backupType == 0.2)
                                LineColor = Color.DeepPink;

                            dbIndex = dbIndex + backupType;

                            var scatter = backupTimelinePlot.AddScatter(
                                new double[] { startDate, finishDate },
                                new double[] { dbIndex, dbIndex },
                                color: LineColor,
                                markerShape: MarkerShape.filledCircle,
                                markerSize: 5,
                                lineWidth: 2);
                            scatter.MarkerLineWidth = 1;

                        }
                    }
                }
            }

            backupTimelinePlot.YLabel("Backup Duration");
            backupTimelinePlot.Axes.Left.SetTicks(customLabels.Keys.ToArray(), customLabels.Values.ToArray());
            backupTimelinePlot.SetAxisLimits(yMin: -0.5, yMax: customLabels.Count + 0.5);

            BackupTimelineModel.Refresh();


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

            var piePlot = BackupSizeModel.Plot;
            piePlot.Clear();
            piePlot.Title("Database Backup Sizes");

            var pieLabels = new List<string>();
            var pieSizes = new List<double>();

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

                            pieLabels.Add(databaseName);
                            pieSizes.Add(backupSize);

                        }
                    }
                }
            }
            if (pieSizes.Count > 0)
            {
                var pie = piePlot.AddPie(pieSizes.ToArray());
                pie.SliceLabels = pieLabels.ToArray();
                pie.Explode = true;
                piePlot.ShowLegend();
            }

            BackupSizeModel.Refresh();
            */

        }

        private void AgentJobsTimelineModelRefresh_Click(object sender, RoutedEventArgs e)
        {
            /*
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

            / * CREATE INDEX IDX_sysjobhistory_1
                ON sysjobhistory(job_id, step_id, run_date, run_time) 
	            WITH (DATA_COMPRESSION = PAGE, ONLINE = ON, MAXDOP = 4); * /
            ";

            var jobsPlot = AgentJobsTimelineModel.Plot;
            jobsPlot.Clear();
            jobsPlot.Title("Agent Jobs: Frequency and Durations Analysis");
            jobsPlot.XAxis.DateTimeFormat(true);
            jobsPlot.XLabel("Date");

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

                            var LineColor = Color.DeepPink;
                            if (resultType == 0) // Failure
                                LineColor = Color.Red;
                            else if (resultType == 1) // Success
                                LineColor = Color.Green;
                            else if (resultType == 2) // ??
                                LineColor = Color.DeepPink;
                            else if (resultType == 3) // Stopped manually
                                LineColor = Color.Purple;

                            var scatter = jobsPlot.AddScatter(
                                new double[] { startDate, finishDate },
                                new double[] { jIndex, jIndex },
                                color: LineColor,
                                markerShape: MarkerShape.filledCircle,
                                markerSize: 5,
                                lineWidth: 2);
                            scatter.MarkerLineWidth = 1;

                        }
                    }
                }
            }

            jobsPlot.YLabel("Job Duration");
            jobsPlot.Axes.Left.SetTicks(customLabels.Keys.ToArray(), customLabels.Values.ToArray());
            jobsPlot.SetAxisLimits(yMin: -0.5, yMax: customLabels.Count + 0.5);

            AgentJobsTimelineModel.Refresh();
            */

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
            if (!string.IsNullOrEmpty(_newVersionURL))
            {
                Process.Start(new ProcessStartInfo(_newVersionURL) { UseShellExecute = true });
            } 
        }

    }
}