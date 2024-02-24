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

    /// <summary>
    /// Interaction logic for HealthDashboard_ServerControl.
    /// </summary>
    public partial class HealthDashboard_ServerControl : UserControl
    {

        public string connectionString = null;
        public DateTime lastRefresh = new DateTime(2000, 1, 1, 0, 0, 0);
        private CancellationTokenSource _cancellationTokenSource;
        private bool _disposed = false;
        private bool _monitoringStarted = false;

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



            UpdateUI(0, new HealthDashboardServerMetric { }, true);

            if (ServiceCache.ScriptFactory.CurrentlyActiveWndConnectionInfo != null)
            {

                var connectionInfo = ScriptFactoryAccess.GetCurrentConnectionInfo(inMaster: true);
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
            //string serverMemory = FormatBytesToGB(0 * 1024);

            Label_MemoryInfo.Content = $"{usedMemory} / {targetMemory}";
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
                                MarkerType = MarkerType.Circle,
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
                                MarkerType = MarkerType.Circle,
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

    }
}