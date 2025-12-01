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
    using OxyPlot;
    using OxyPlot.Axes;
    using OxyPlot.Series;
    using System.Collections.Generic;
    using System.Windows.Input;
    using OxyPlot.Legends;
    using System.Linq;
    using static HealthDashboardServerMetric;
    using System.Diagnostics;
    using static AxialSqlTools.AxialSqlToolsPackage;
    using DocumentFormat.OpenXml.Bibliography;
    using DocumentFormat.OpenXml.Spreadsheet;
    using static AxialSqlTools.HealthDashboard_ServerControl;
    using MarkerType = OxyPlot.MarkerType;

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

            //--------------------------------------------------------------------
            // Disk info graph
            var barModel = new PlotModel { Title = "Volume(s) Utilization" };

            var barSeries1 = new BarSeries
            {
                LabelPlacement = LabelPlacement.Inside,
                LabelFormatString = "{0:0} Gb", // Adjust this to change how the labels are formatted
                StrokeColor = OxyColors.Black,
                StrokeThickness = 1,
                IsStacked = true
            };
            foreach (var disk in metrics.DisksInfo)
                barSeries1.Items.Add(new BarItem { Value = disk.UsedSpaceGb, Color = OxyColors.LightPink });
            barModel.Series.Add(barSeries1);

            var barSeries2 = new BarSeries
            {
                LabelPlacement = LabelPlacement.Inside,
                LabelFormatString = "{0:0} Gb", // Adjust this to change how the labels are formatted
                StrokeColor = OxyColors.Black,
                StrokeThickness = 1,
                IsStacked = true
            };
            foreach (var disk in metrics.DisksInfo)
                barSeries2.Items.Add(new BarItem { Value = disk.FreeSpaceGb, Color = OxyColors.LightBlue });
            barModel.Series.Add(barSeries2);


            barModel.Axes.Add(new CategoryAxis
            {
                Position = AxisPosition.Left,
                Key = "DiskAxis",
                ItemsSource = metrics.DisksInfo.Select(disk => disk.VolumeDescription).ToList(),
                IsZoomEnabled = false,
                IsPanEnabled = false
            });

            barModel.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Bottom,
                MinimumPadding = 0.1,
                MaximumPadding = 0.1,
                AbsoluteMinimum = 0,
                Title = "Gb",
                IsZoomEnabled = false,
                IsPanEnabled = false
            });

            this.DiskInfoModel.Model = barModel;


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

        private void UpdateWaitStatsGraph(List<WaitsInfo> previousWaitStats, Dictionary<DateTime, List<WaitsInfo>> aggrData)
        {

            var sortedKeys = aggrData.Keys.OrderBy(k => k).ToList();

            var barModelWS = new PlotModel { Title = "Real-time Wait Stats" };

            var categoryAxis = new CategoryAxis { 
                Position = AxisPosition.Left,
                IsZoomEnabled = false,
                IsPanEnabled = false
            };
            var valueAxis = new LinearAxis { 
                Position = AxisPosition.Bottom, 
                MinimumPadding = 0, 
                AbsoluteMinimum = 0,
                IsZoomEnabled = false,
                IsPanEnabled = false
            };

            barModelWS.Axes.Add(categoryAxis);
            barModelWS.Axes.Add(valueAxis);

            foreach (var key in sortedKeys)
            {
                categoryAxis.Labels.Add(key.ToString("HH:mm"));
            }

            var LegendWS = new Legend
            {
                LegendTitle = "Wait Stats",
                LegendPosition = LegendPosition.RightTop,
                LegendPlacement = LegendPlacement.Outside,
                LegendOrientation = LegendOrientation.Vertical,
                LegendBackground = OxyColor.FromAColor(200, OxyColors.White),
                LegendBorder = OxyColors.Black
            };
            LegendWS.LegendMaxWidth = 200;

            // Legend configuration
            barModelWS.IsLegendVisible = true; // Make the legend visible
            barModelWS.Legends.Add(LegendWS);

            foreach (var previousWaitStat in previousWaitStats)
            {
                var barSeriesWS = new BarSeries
                {
                    //LabelPlacement = LabelPlacement.Inside,
                    //LabelFormatString = "{0:0}", // Adjust this to change how the labels are formatted
                    StrokeColor = OxyColors.Black,
                    StrokeThickness = 1,
                    IsStacked = true
                };

                barSeriesWS.Title = previousWaitStat.WaitName;

                foreach (var aggValue in aggrData)
                {
                    foreach (WaitsInfo ws in aggValue.Value)
                        if (ws.WaitName == previousWaitStat.WaitName)
                            barSeriesWS.Items.Add(new BarItem { Value = (double)ws.WaitSec }); //, Color = OxyColors.LightPink });
                }
                barModelWS.Series.Add(barSeriesWS);
            }

            this.WaitStatsModel.Model = barModelWS;
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

        private PlotModel CreateTimeSeriesModel(string title, string yAxisTitle, Func<PerformanceSample, double> valueSelector, bool clampToZero = true)
        {
            var model = new PlotModel { Title = title };

            model.Axes.Add(new DateTimeAxis
            {
                Position = AxisPosition.Bottom,
                StringFormat = "HH:mm",
                IntervalType = DateTimeIntervalType.Minutes,
                IsZoomEnabled = false,
                IsPanEnabled = false,
                MinorIntervalType = DateTimeIntervalType.Minutes
            });

            var linearAxis = new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = yAxisTitle,
                IsZoomEnabled = false,
                IsPanEnabled = false
            };

            if (clampToZero)
            {
                linearAxis.Minimum = 0;
            }

            model.Axes.Add(linearAxis);

            var series = new LineSeries
            {
                StrokeThickness = 2,
                MarkerSize = 2,
                MarkerType = MarkerType.Circle
            };

            foreach (var sample in _performanceSamples.OrderBy(s => s.Timestamp))
            {
                series.Points.Add(DateTimeAxis.CreateDataPoint(sample.Timestamp, valueSelector(sample)));
            }

            model.Series.Add(series);

            return model;
        }

        private void UpdatePerformanceCharts()
        {
            PerfChart_CpuUtilization.Model = CreateTimeSeriesModel("CPU Utilization (%)", "%", s => s.CpuUtilization);
            PerfChart_UserConnections.Model = CreateTimeSeriesModel("User Connections", "connections", s => s.UserConnections);
            PerfChart_BatchRequests.Model = CreateTimeSeriesModel("Batch Requests/sec", "requests/sec", s => s.BatchRequestsSec);
            PerfChart_SqlCompilations.Model = CreateTimeSeriesModel("SQL Compilations/sec", "compilations/sec", s => s.SqlCompilationsSec);
            PerfChart_PageLifeExpectancy.Model = CreateTimeSeriesModel("Page Life Expectancy", "seconds", s => s.PageLifeExpectancy, clampToZero: false);
            PerfChart_PageReads.Model = CreateTimeSeriesModel("Page Reads/sec", "pages/sec", s => s.PageReadsSec);
            PerfChart_PageWrites.Model = CreateTimeSeriesModel("Page Writes/sec", "pages/sec", s => s.PageWritesSec);
            PerfChart_LogFlushes.Model = CreateTimeSeriesModel("Log Flushes/sec", "flushes/sec", s => s.LogFlushesSec);
            PerfChart_Transactions.Model = CreateTimeSeriesModel("Transactions/sec", "transactions/sec", s => s.TransactionsSec);
            PerfChart_LockWaits.Model = CreateTimeSeriesModel("Lock Waits/sec", "waits/sec", s => s.LockWaitsSec);
            PerfChart_MemoryGrantsPending.Model = CreateTimeSeriesModel("Memory Grants Pending", "grants", s => s.MemoryGrantsPending);
            PerfChart_TotalServerMemory.Model = CreateTimeSeriesModel("Total Server Memory", "MB", s => s.TotalServerMemoryMb);
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

            var MyModel = new PlotModel { Title = "Database Backups: Frequency and Durations Analysis" };

            MyModel.Axes.Add(new DateTimeAxis
            {
                Position = AxisPosition.Bottom,
                StringFormat = "dd-MM-yyyy HH:mm",
                Title = "Date",
                IntervalType = DateTimeIntervalType.Hours,
                MinorIntervalType = DateTimeIntervalType.Minutes,
                IntervalLength = 80,
                IsZoomEnabled = false,
                IsPanEnabled = false
            });

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


                            var startDate = DateTimeAxis.ToDouble(reader.GetDateTime(1));
                            var finishDate = DateTimeAxis.ToDouble(reader.GetDateTime(2));

                            var backupType = (double)reader.GetDecimal(3);

                            var LineColor = OxyColors.Blue;
                            if (backupType == 0.1)
                                LineColor = OxyColors.Green;
                            if (backupType == 0.2)
                                LineColor = OxyColors.DeepPink;

                            dbIndex = dbIndex + backupType;

                            var scatterSeries = new ScatterSeries {
                                MarkerType = OxyPlot.MarkerType.Circle,
                                MarkerFill = LineColor,
                                MarkerStrokeThickness = 1
                            };

                            // Add two points to the scatter series
                            scatterSeries.Points.Add(new ScatterPoint(startDate, dbIndex));
                            scatterSeries.Points.Add(new ScatterPoint(finishDate, dbIndex));

                            // Add a line series to connect the dots
                            var lineSeries = new LineSeries()
                            {
                                Color = LineColor,
                                StrokeThickness = 2,
                                LineStyle = LineStyle.Solid
                            };
                            lineSeries.Points.Add(new DataPoint(startDate, dbIndex));
                            lineSeries.Points.Add(new DataPoint(finishDate, dbIndex));

                            MyModel.Series.Add(lineSeries);
                            MyModel.Series.Add(scatterSeries);

                        }
                    }
                }
            }

            var yAxis = new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "Backup Duration",
                MajorStep = 1,
                MinorStep = 1,
                IsZoomEnabled = false,
                IsPanEnabled = false,
                
                LabelFormatter = value =>
                {
                    // Return the custom label if it exists; otherwise, return the default string representation
                    if (customLabels.TryGetValue(value, out var label))
                    {
                        return label;
                    }
                    return value.ToString();
                }
            };

            MyModel.Axes.Add(yAxis);


            this.BackupTimelineModel.Model = MyModel;


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

            var PieModel = new PlotModel { Title = "Database Backup Sizes"};

            dynamic seriesP1 = new PieSeries { StrokeThickness = 2.0, InsideLabelPosition = 0.8, AngleSpan = 360, StartAngle = 0 };

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

                            seriesP1.Slices.Add(new PieSlice(databaseName, backupSize) { IsExploded = true });

                        }
                    }
                }
            }


            PieModel.Series.Add(seriesP1);

            this.BackupSizeModel.Model = PieModel;

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

            var MyModel = new PlotModel { Title = "Agent Jobs: Frequency and Durations Analysis" };

            MyModel.Axes.Add(new DateTimeAxis
            {
                Position = AxisPosition.Bottom,
                StringFormat = "dd-MM-yyyy HH:mm",
                Title = "Date",
                IntervalType = DateTimeIntervalType.Hours,
                MinorIntervalType = DateTimeIntervalType.Minutes,
                IntervalLength = 80,
                IsZoomEnabled = false,
                IsPanEnabled = false
            });

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


                            var startDate = DateTimeAxis.ToDouble(reader.GetDateTime(1));
                            var finishDate = DateTimeAxis.ToDouble(reader.GetDateTime(2));

                            var resultType = reader.GetInt32(3);

                            var LineColor = OxyColors.DeepPink;
                            if (resultType == 0) // Failure
                                LineColor = OxyColors.Red;
                            else if (resultType == 1) // Success
                                LineColor = OxyColors.Green;
                            else if (resultType == 2) // ??
                                LineColor = OxyColors.DeepPink;
                            else if (resultType == 3) // Stopped manually
                                LineColor = OxyColors.Purple;
                            
                            var scatterSeries = new ScatterSeries
                            {
                                MarkerType = OxyPlot.MarkerType.Circle,
                                MarkerFill = LineColor,
                                MarkerStrokeThickness = 1
                            };

                            // Add two points to the scatter series
                            scatterSeries.Points.Add(new ScatterPoint(startDate, jIndex));
                            scatterSeries.Points.Add(new ScatterPoint(finishDate, jIndex));

                            // Add a line series to connect the dots
                            var lineSeries = new LineSeries()
                            {
                                Color = LineColor,
                                StrokeThickness = 2,
                                LineStyle = LineStyle.Solid
                            };
                            lineSeries.Points.Add(new DataPoint(startDate, jIndex));
                            lineSeries.Points.Add(new DataPoint(finishDate, jIndex));

                            MyModel.Series.Add(lineSeries);
                            MyModel.Series.Add(scatterSeries);

                        }
                    }
                }
            }

            var yAxis = new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "Job Duration",
                MajorStep = 1,
                MinorStep = 1,
                IsZoomEnabled = false,
                IsPanEnabled = false,
                LabelFormatter = value =>
                {
                    // Return the custom label if it exists; otherwise, return the default string representation
                    if (customLabels.TryGetValue(value, out var label))
                    {
                        return label;
                    }
                    return value.ToString();
                }
            };

            MyModel.Axes.Add(yAxis);

            this.AgentJobsTimelineModel.Model = MyModel;


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