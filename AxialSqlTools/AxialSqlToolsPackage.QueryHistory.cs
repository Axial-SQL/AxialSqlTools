using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System;
using System.Collections.Concurrent;
using System.IO;
using Task = System.Threading.Tasks.Task;

namespace AxialSqlTools
{
    // Query history queue and persistence to SQL Server or JSONL files.
    public sealed partial class AxialSqlToolsPackage
    {
        private class QueryHistoryEntry
        {
            public DateTime StartTime;
            public DateTime FinishTime;
            public string ElapsedTime;
            public long TotalRowsReturned;
            public string ExecResult;
            public string QueryText;
            public string DataSource;
            public string DatabaseName;
            public string LoginName;
            public string WorkstationId;
        }

        private const string QueryHistoryStorageModeTextFiles = "TextFiles";

        private const string QueryHistoryStorageModeDisabled = "Disabled";

        private static ConcurrentQueue<QueryHistoryEntry> _queryHistoryQueue = new ConcurrentQueue<QueryHistoryEntry>();

        private static void EnqueueDataForProcessing(QueryHistoryEntry data)
        {
            _queryHistoryQueue.Enqueue(data);
            _ = Task.Run(() => ProcessDataAsync());
        }

        private static async Task ProcessDataAsync()
        {
            while (_queryHistoryQueue.TryDequeue(out QueryHistoryEntry data))
            {
                await PersistDataAsync(data);
            }

        }

        private static async Task PersistDataAsync(QueryHistoryEntry data)
        {
            try
            {
                string storageMode = SettingsManager.GetQueryHistoryStorageMode();
                if (string.Equals(storageMode, QueryHistoryStorageModeDisabled, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                if (string.Equals(storageMode, QueryHistoryStorageModeTextFiles, StringComparison.OrdinalIgnoreCase))
                {
                    await PersistDataAsJsonLineAsync(data);
                    return;
                }

                string connectionString = SettingsManager.GetQueryHistoryConnectionString();
                string qhTableName = SettingsManager.GetQueryHistoryTableNameOrDefault();
                string indexNameGuid = Guid.NewGuid().ToString(); // too much complexity trying to incorporate all possible table name combinations into proper index name

                if (string.IsNullOrEmpty(connectionString))
                {
                    return;
                }

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();
                    string sql = $@"
                        IF OBJECT_ID('{qhTableName}') IS NULL
                        BEGIN
                            CREATE TABLE {qhTableName} (
                                [QueryID]           INT            IDENTITY (1, 1) NOT NULL,
                                [StartTime]         DATETIME       NOT NULL,
                                [FinishTime]        DATETIME       NOT NULL,
                                [ElapsedTime]       VARCHAR (15)   NOT NULL,
                                [TotalRowsReturned] BIGINT         NOT NULL,
                                [ExecResult]        VARCHAR (100)  NOT NULL,
                                [QueryText]         NVARCHAR (MAX) NOT NULL,
                                [DataSource]        NVARCHAR (128) NOT NULL,
                                [DatabaseName]      NVARCHAR (128) NOT NULL,
                                [LoginName]         NVARCHAR (128) NOT NULL,
                                [WorkstationId]     NVARCHAR (128) NOT NULL,
                                PRIMARY KEY CLUSTERED ([QueryID]),
                                INDEX [IDX_{indexNameGuid}_1] ([StartTime]),
                                INDEX [IDX_{indexNameGuid}_2] ([FinishTime]),
                                INDEX [IDX_{indexNameGuid}_3] ([DataSource]),
                                INDEX [IDX_{indexNameGuid}_4] ([DatabaseName])
                            );
                            ALTER INDEX ALL ON {qhTableName} REBUILD WITH (DATA_COMPRESSION = PAGE);
                        END

                        INSERT INTO {qhTableName}
                            (StartTime, FinishTime, ElapsedTime, TotalRowsReturned, 
                                ExecResult, QueryText, DataSource, DatabaseName, LoginName, WorkstationId) 
                        VALUES (@StartTime, @FinishTime, @ElapsedTime, @TotalRowsReturned, 
                                    @ExecResult, @QueryText, @DataSource, @DatabaseName, @LoginName, @WorkstationId)
                        ";

                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@StartTime", data.StartTime);
                        command.Parameters.AddWithValue("@FinishTime", data.FinishTime);
                        command.Parameters.AddWithValue("@ElapsedTime", data.ElapsedTime);
                        command.Parameters.AddWithValue("@TotalRowsReturned", data.TotalRowsReturned);
                        command.Parameters.AddWithValue("@ExecResult", data.ExecResult);
                        command.Parameters.AddWithValue("@QueryText", data.QueryText?.Trim() ?? string.Empty);
                        command.Parameters.AddWithValue("@DataSource", data.DataSource ?? string.Empty);
                        command.Parameters.AddWithValue("@DatabaseName", data.DatabaseName ?? string.Empty);
                        command.Parameters.AddWithValue("@LoginName", data.LoginName ?? string.Empty);
                        command.Parameters.AddWithValue("@WorkstationId", data.WorkstationId ?? string.Empty);
                        await command.ExecuteNonQueryAsync();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "[QueryHistory-PersistDataAsync]: An exception occurred");
            }

        }

        private static Task PersistDataAsJsonLineAsync(QueryHistoryEntry data)
        {
            string folderPath = SettingsManager.GetQueryHistoryTextFileFolder();
            Directory.CreateDirectory(folderPath);

            string fileName = $"query-history-{DateTime.UtcNow:yyyy-MM-dd}.jsonl";
            string filePath = Path.Combine(folderPath, fileName);
            string json = JsonConvert.SerializeObject(data);

            File.AppendAllText(filePath, json + Environment.NewLine);
            return Task.CompletedTask;
        }
    }
}
