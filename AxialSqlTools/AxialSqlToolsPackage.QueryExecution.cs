using EnvDTE;
using Microsoft.Data.SqlClient;
using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace AxialSqlTools
{
    // Query execution hooks, result display updates, and statistics capture.
    public sealed partial class AxialSqlToolsPackage
    {
        private static int _statisticsCaptureVersion;

        private static int _pendingStatisticsCaptureVersion;

        private static readonly object _statisticsCaptureSyncRoot = new object();

        private static CancellationTokenSource _statisticsCaptureCancellationTokenSource;

        public CommandEvents m_queryExecuteEvent { get; private set; }

        private static void AttachStatisticsExecutionCompletedHandler(object sqlResultsControl)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (sqlResultsControl == null)
            {
                return;
            }

            EventHandler eventHandler = SQLResultsControl_ScriptExecutionCompleted;

            EventInfo eventInfo = sqlResultsControl.GetType().GetEvent("ScriptExecutionCompleted");
            if (eventInfo == null)
            {
                return;
            }

            Delegate handlerDelegate = Delegate.CreateDelegate(eventInfo.EventHandlerType, eventHandler.Target, eventHandler.Method);
            eventInfo.RemoveEventHandler(sqlResultsControl, handlerDelegate);
            eventInfo.AddEventHandler(sqlResultsControl, handlerDelegate);
        }

        public static void EnsureStatisticsExecutionHookForActiveWindow(string reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                var sqlResultsControl = GridAccess.GetSQLResultsControl();
                AttachStatisticsExecutionCompletedHandler(sqlResultsControl);
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, $"Failed to ensure statistics execution hook ({reason}).");
            }
        }

        //----------------


        // This method aligns all numeric values to the right
        public static void SQLResultsControl_ScriptExecutionCompleted(object QEOLESQLExec, object b)
        {

            ThreadHelper.ThrowIfNotOnUIThread();

            var generalSettings = SettingsManager.GetGeneralSettings();

            try
            {
                if (generalSettings.alignNumericValuesToRight)
                {
                    //1. Align numeric types to the right
                    CollectionBase gridContainers = GridAccess.GetGridContainers();

                    foreach (var gridContainer in gridContainers)
                    {
                        var grid = GridAccess.GetNonPublicField(gridContainer, "m_grid") as GridControl;
                        var gridStorage = grid.GridStorage;
                        var schemaTable = GridAccess.GetNonPublicField(gridStorage, "m_schemaTable") as DataTable;

                        var gridColumns = GridAccess.GetNonPublicField(grid, "m_Columns") as GridColumnCollection;
                        if (gridColumns != null)
                        {
                            //Why no "flot"? Because it cannot be aligned "good" due to the varying number of digits in the decimal part.
                            string[] typeToAlignRight = new string[] { "tinyint", "smallint", "int", "bigint", "money", "smallmoney", "decimal", "numeric" };

                            List<int> columnsToAlignRight = new List<int> { };

                            for (int c = 0; c < schemaTable.Rows.Count; c++)
                            {
                                int columnOrdinal = (int)schemaTable.Rows[c][1];
                                var sqlDataTypeName = schemaTable.Rows[c][24];

                                if (typeToAlignRight.Contains(sqlDataTypeName))
                                {
                                    columnsToAlignRight.Add(columnOrdinal);
                                }
                            }

                            foreach (Microsoft.SqlServer.Management.UI.Grid.GridColumn gridColumn in gridColumns)
                            {

                                if (columnsToAlignRight.Contains(gridColumn.ColumnIndex - 1) || gridColumn.ColumnIndex == 0)
                                {
                                    // not needed
                                    //var textAlignField = GridAccess.GetNonPublicFieldInfo(gridColumn, "TextAlign");
                                    //if (textAlignField != null)
                                    //{
                                    //    textAlignField.SetValue(gridColumn, System.Windows.Forms.HorizontalAlignment.Right);
                                    //}

                                    // applies to the row number column 
                                    var textAlignField2 = GridAccess.GetNonPublicFieldInfo(gridColumn, "m_myAlign");
                                    if (textAlignField2 != null)
                                    {
                                        textAlignField2.SetValue(gridColumn, System.Windows.Forms.HorizontalAlignment.Right);
                                    }

                                    var textAlignField3 = GridAccess.GetNonPublicFieldInfo(gridColumn, "m_textFormat");
                                    if (textAlignField3 != null)
                                    {
                                        System.Windows.Forms.TextFormatFlags flags = (System.Windows.Forms.TextFormatFlags)GridAccess.GetNonPublicField(gridColumn, "m_textFormat");
                                        textAlignField3.SetValue(gridColumn, flags | System.Windows.Forms.TextFormatFlags.Right);
                                    }
                                }

                            }
                        }

                        grid.Refresh();

                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "An exception occurred");
            }

            try
            {
                // 2. Get open transaction info
                int openTranCount = 0;

                var SQLResultsControl = GridAccess.GetSQLResultsControl();
                var m_SqlExec = GridAccess.GetNonPublicField(SQLResultsControl, "m_sqlExec");

                Microsoft.Data.SqlClient.SqlConnection connection = GridAccess.GetNonPublicField(m_SqlExec, "m_conn") as Microsoft.Data.SqlClient.SqlConnection;
                if (generalSettings.useTransactionWarning && connection != null && connection.State == ConnectionState.Open)
                {
                    using (Microsoft.Data.SqlClient.SqlCommand command = new Microsoft.Data.SqlClient.SqlCommand("SELECT @@TRANCOUNT", connection))
                    {
                        var result = command.ExecuteScalar();
                        openTranCount = result != DBNull.Value ? Convert.ToInt32(result) : 0;
                    }
                }

                bool isColumnEncryptionSettingOn = generalSettings.useAlwaysEncryptedWarning && connection != null
                    && new SqlConnectionStringBuilder(connection.ConnectionString).ColumnEncryptionSetting == SqlConnectionColumnEncryptionSetting.Enabled;

                string elapsedTime = null;
                if (generalSettings.usePreciseExecutionTime)
                {
                    var editorProperties = GridAccess.GetNonPublicField(m_SqlExec, "editorProperties");
                    elapsedTime = (string)GridAccess.GetProperty(editorProperties, "ElapsedTime");
                }

                GridAccess.ChangeStatusBarContent(openTranCount, isColumnEncryptionSettingOn, elapsedTime, generalSettings);

                // Re-apply connection color after status bar update
                string dataSource = (string)GridAccess.GetProperty(connection, "DataSource");
                string database = (string)GridAccess.GetProperty(connection, "Database");
                GridAccess.ApplyConnectionColor(dataSource, database);

            }
            catch (Exception ex)
            {
                _logger.Error(ex, "An exception occurred");
            }

            // query history
            try
            {

                var editorProperties = GridAccess.GetNonPublicField(QEOLESQLExec, "editorProperties");

                var textSpan = GridAccess.GetNonPublicField(QEOLESQLExec, "textSpan");

                var mConn = GridAccess.GetNonPublicField(QEOLESQLExec, "m_conn");

                var QueryHistoryObj = new QueryHistoryEntry();
                QueryHistoryObj.StartTime = (DateTime)GridAccess.GetNonPublicField(editorProperties, "startTime");
                QueryHistoryObj.FinishTime = (DateTime)GridAccess.GetNonPublicField(editorProperties, "finishTime");
                QueryHistoryObj.ElapsedTime = (string)GridAccess.GetProperty(editorProperties, "ElapsedTime");
                QueryHistoryObj.TotalRowsReturned = (long)GridAccess.GetProperty(editorProperties, "TotalRowsReturned");

                // Success, Failure -> not really clear how to track batch execution results...
                QueryHistoryObj.ExecResult = GridAccess.GetNonPublicField(QEOLESQLExec, "m_execResult").ToString();

                QueryHistoryObj.QueryText = (string)GridAccess.GetProperty(textSpan, "Text");
                QueryHistoryObj.DataSource = (string)GridAccess.GetProperty(mConn, "DataSource");
                QueryHistoryObj.DatabaseName = (string)GridAccess.GetProperty(mConn, "Database");
                QueryHistoryObj.LoginName = (string)GridAccess.GetProperty(editorProperties, "ChildLoginName");
                QueryHistoryObj.WorkstationId = (string)GridAccess.GetProperty(mConn, "WorkstationId");

                EnqueueDataForProcessing(QueryHistoryObj);

            }
            catch (Exception ex)
            {
                _logger.Error(ex, "An exception occurred");
            }

            if (!StatisticsSummaryStore.IsWindowOpen())
            {
                ThreadHelper.Generic.BeginInvoke(() => EnsureStatisticsExecutionHookForActiveWindow("post-skip"));

                return;
            }

            var captureVersion = Interlocked.Exchange(ref _pendingStatisticsCaptureVersion, 0);
            if (captureVersion == 0)
            {
                captureVersion = Interlocked.Increment(ref _statisticsCaptureVersion);
            }

            if (!StatisticsSummaryStore.BeginCapture(captureVersion))
            {
                return;
            }

            var captureCancellationTokenSource = CreateStatisticsCaptureCancellationTokenSource();

            _ = Task.Run(async delegate
            {
                try
                {
                    await CaptureStatisticsSummaryAsync(QEOLESQLExec, captureVersion, captureCancellationTokenSource.Token);
                }
                catch (OperationCanceledException)
                {
                }
                catch (Exception ex)
                {
                    StatisticsSummaryStore.MarkUnavailable(captureVersion);
                    _logger.Error(ex, "Failed to capture statistics summary.");
                }
                finally
                {
                    ReleaseStatisticsCaptureCancellationTokenSource(captureCancellationTokenSource);
                }
            });


        }

        public static void CancelStatisticsCapture(bool updateStore = true)
        {
            CancellationTokenSource captureCancellationTokenSource;

            Interlocked.Exchange(ref _pendingStatisticsCaptureVersion, 0);

            lock (_statisticsCaptureSyncRoot)
            {
                captureCancellationTokenSource = _statisticsCaptureCancellationTokenSource;
                _statisticsCaptureCancellationTokenSource = null;
            }

            captureCancellationTokenSource?.Cancel();

            if (updateStore)
            {
                StatisticsSummaryStore.CancelCapture();
            }
        }

        private static CancellationTokenSource CreateStatisticsCaptureCancellationTokenSource()
        {
            CancellationTokenSource previousCancellationTokenSource;
            CancellationTokenSource nextCancellationTokenSource;

            lock (_statisticsCaptureSyncRoot)
            {
                previousCancellationTokenSource = _statisticsCaptureCancellationTokenSource;
                nextCancellationTokenSource = new CancellationTokenSource();
                _statisticsCaptureCancellationTokenSource = nextCancellationTokenSource;
            }

            previousCancellationTokenSource?.Cancel();
            return nextCancellationTokenSource;
        }

        private static void ReleaseStatisticsCaptureCancellationTokenSource(CancellationTokenSource captureCancellationTokenSource)
        {
            lock (_statisticsCaptureSyncRoot)
            {
                if (ReferenceEquals(_statisticsCaptureCancellationTokenSource, captureCancellationTokenSource))
                {
                    _statisticsCaptureCancellationTokenSource = null;
                }
            }

            captureCancellationTokenSource.Dispose();
        }

        private static async Task CaptureStatisticsSummaryAsync(object sqlExecutionContext, int captureVersion, CancellationToken cancellationToken)
        {
            const int maxAttempts = 3;

            cancellationToken.ThrowIfCancellationRequested();
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var statisticsOutputOptions = GetStatisticsSummaryOutputOptions(sqlExecutionContext);
            if (!statisticsOutputOptions.HasStatisticsOutput)
            {
                StatisticsSummaryStore.MarkUnavailable(captureVersion, StatisticsSummaryCaptureStatus.StatisticsDisabled);
                return;
            }

            for (int attempt = 0; attempt < maxAttempts; attempt++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!StatisticsSummaryStore.IsWindowOpen())
                {
                    throw new OperationCanceledException(cancellationToken);
                }

                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

                GridAccess.TryFlushStatisticsMessages();

                var statisticsText = GridAccess.TryGetStatisticsMessagesText();
                if (TryStoreStatisticsSummary(sqlExecutionContext, statisticsText, captureVersion, out var summary))
                {
                    return;
                }

                if (attempt < maxAttempts - 1)
                {
                    await Task.Delay(150, cancellationToken);
                }
            }

            StatisticsSummaryStore.MarkUnavailable(captureVersion);
        }

        private sealed class StatisticsSummaryOutputOptions
        {
            public bool StatisticsIoEnabled { get; set; }
            public bool StatisticsTimeEnabled { get; set; }
            public bool HasStatisticsOutput => StatisticsIoEnabled || StatisticsTimeEnabled;
        }

        private static StatisticsSummaryOutputOptions GetStatisticsSummaryOutputOptions(object sqlExecutionContext)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                if (!(GridAccess.GetNonPublicField(sqlExecutionContext, "m_conn") is SqlConnection connection)
                    || connection.State != ConnectionState.Open)
                {
                    return new StatisticsSummaryOutputOptions
                    {
                        StatisticsIoEnabled = true,
                        StatisticsTimeEnabled = true,
                    };
                }

                using (var command = new SqlCommand("DBCC USEROPTIONS WITH NO_INFOMSGS;", connection))
                using (var reader = command.ExecuteReader())
                {
                    var options = new StatisticsSummaryOutputOptions();

                    while (reader.Read())
                    {
                        if (reader.FieldCount < 2)
                        {
                            continue;
                        }

                        var optionName = reader.IsDBNull(0) ? null : reader.GetString(0);
                        var optionValue = reader.IsDBNull(1) ? null : reader.GetString(1);
                        if (string.IsNullOrWhiteSpace(optionName))
                        {
                            continue;
                        }

                        if (optionName.Equals("statistics io", StringComparison.OrdinalIgnoreCase))
                        {
                            options.StatisticsIoEnabled = IsUserOptionEnabled(optionValue);
                        }
                        else if (optionName.Equals("statistics time", StringComparison.OrdinalIgnoreCase))
                        {
                            options.StatisticsTimeEnabled = IsUserOptionEnabled(optionValue);
                        }
                    }

                    return options;
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to inspect session statistics options.");
                return new StatisticsSummaryOutputOptions
                {
                    StatisticsIoEnabled = true,
                    StatisticsTimeEnabled = true,
                };
            }
        }

        private static bool IsUserOptionEnabled(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            return value.Equals("on", StringComparison.OrdinalIgnoreCase)
                || value.Equals("set", StringComparison.OrdinalIgnoreCase)
                || value.Equals("true", StringComparison.OrdinalIgnoreCase)
                || value.Equals("1", StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryStoreStatisticsSummary(object sqlExecutionContext, string statisticsText, int captureVersion, out StatisticsSummary summary)
        {
            summary = null;

            if (string.IsNullOrWhiteSpace(statisticsText))
            {
                return false;
            }

            summary = StatisticsSummaryParser.Parse(statisticsText);
            if (summary == null)
            {
                return false;
            }

            var textSpan = GridAccess.GetNonPublicField(sqlExecutionContext, "textSpan");
            var mConn = GridAccess.GetNonPublicField(sqlExecutionContext, "m_conn");

            summary.QueryText = GridAccess.GetProperty(textSpan, "Text") as string;
            summary.DataSource = GridAccess.GetProperty(mConn, "DataSource") as string;
            summary.DatabaseName = GridAccess.GetProperty(mConn, "Database") as string;

            StatisticsSummaryStore.Set(summary, captureVersion);
            return true;
        }

        private void CommandEvents_BeforeExecute(string Guid, int ID, object CustomIn, object CustomOut, ref bool CancelDefault)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                if (CancelDefault) return;
                try
                {
                    if (QuerySafety.FatalActionGuard.ShouldCancel(GetGlobalService(typeof(DTE)) as DTE))
                    {
                        CancelDefault = true;
                        return;
                    }
                }
                catch (Exception ex)
                {
                    CancelDefault = true;
                    _logger?.Error(ex, "Fatal action check failed. Query execution cancelled.");
                    return;
                }

                EnsureStatisticsExecutionHookForActiveWindow("query-execute-before");

                if (!StatisticsSummaryStore.IsWindowOpen())
                {
                    return;
                }

                var captureVersion = Interlocked.Increment(ref _statisticsCaptureVersion);
                Interlocked.Exchange(ref _pendingStatisticsCaptureVersion, captureVersion);

                StatisticsSummaryStore.BeginCapture(captureVersion);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to prepare statistics capture before query execution.");
            }
        }

        //it has been executed, but the Grid hasn't been created yet...
        private void CommandEvents_AfterExecute(string Guid, int ID, object CustomIn, object CustomOut)
        {
            //ThreadHelper.ThrowIfNotOnUIThread();
        }
    }
}
