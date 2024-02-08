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

            
        }

        public void StartMonitoring()
        {
            if (_monitoringStarted) return;

            UpdateUI(0, new HealthDashboardServerMetric { }, true);

            if (ServiceCache.ScriptFactory.CurrentlyActiveWndConnectionInfo != null)
            {

                UIConnectionInfo connection = ServiceCache.ScriptFactory.CurrentlyActiveWndConnectionInfo.UIConnectionInfo;
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = connection.ServerName;
                builder.IntegratedSecurity = string.IsNullOrEmpty(connection.Password);
                builder.Password = connection.Password;
                builder.UserID = connection.UserName;
                builder.InitialCatalog = "master";
                builder.ApplicationName = "Axial SQL Tools";
                connectionString = builder.ToString();

                _cancellationTokenSource = new CancellationTokenSource();

                ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
                {

                    int i = 0;

                    while (!_cancellationTokenSource.IsCancellationRequested)
                    {

                        var metrics = await MetricsService.FetchServerMetricsAsync(connectionString);

                        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(_cancellationTokenSource.Token);

                        UpdateUI(i, metrics, false);

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
                    LabelInternalException.FontWeight = FontWeights.Bold;

                    return;
                }
                else if (!metrics.Completed)
                {
                    return;
                }
            }

            LabelInternalException.Content = "";

            Label_ServerName.Content = metrics.ServerName;
            Label_ServiceName.Content = metrics.ServiceName;

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

            //TODO - atract user attention by blinking
            //userControlOwner.Caption = "! " + DateTime.Now.ToString();

        }

        public static string FormatBytesToMB(long bytes)
        {
            // Convert bytes to gigabytes (GB)
            double megabytes = bytes / 1024.0 / 1024.0; // 1024^3

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

    }
}