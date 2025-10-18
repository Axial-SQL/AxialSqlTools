using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace AxialSqlTools
{


    public class SettingsManager
    {

        public class HealthDashboardServerQueryTexts
        {
            #region m_BlockingRequestsQuery
            private const string m_BlockingRequestsQuery = @"
WITH blocking
AS (SELECT [session_id] AS [spid],
           percent_complete,
           blocked,
           waittime,
           DB_NAME(sp.[dbid]) AS [Database],
           nt_username AS [User],
           er.[status] AS [Status],
           wait_type AS [Wait],
           SUBSTRING(qt.[text], 
				er.statement_start_offset / 2, 
				(CASE 
					WHEN er.statement_end_offset = -1 
						THEN LEN(CONVERT (NVARCHAR (MAX), qt.text)) * 2 
					ELSE er.statement_end_offset 
				END - er.statement_start_offset) / 2) AS [Individual Query],
           qt.[text] AS [Parent Query],
           [program_name] AS Program,
           hostname,
           cpu_time,
           reads,
           start_time,
           sp.login_time,
           sp.last_batch,
           sp.cmd
    FROM sys.dm_exec_requests AS er
         INNER JOIN sys.sysprocesses AS sp ON er.[session_id] = sp.spid 
         CROSS APPLY sys.dm_exec_sql_text(er.[sql_handle]) AS qt
    WHERE [session_id] > 50 AND [session_id] NOT IN (@@SPID))
SELECT * FROM blocking WHERE blocked <> 0
UNION ALL
SELECT * FROM blocking WHERE spid IN (SELECT blocked FROM blocking WHERE blocked <> 0);
";
            #endregion

            #region m_DatabaseLogUsageInfo
            private const string m_DatabaseLogUsageInfo = @"
SELECT db.[name] AS [Database Name],
       SUSER_SNAME(db.owner_sid) AS [Database Owner],
       db.[compatibility_level] AS [DB Compatibility Level],
       db.recovery_model_desc AS [Recovery Model],
       db.log_reuse_wait_desc AS [Log Reuse Wait Description],
       CONVERT (DECIMAL (18, 2), ds.cntr_value / 1024.0) AS [Total Data File Size on Disk (MB)],
       CONVERT (DECIMAL (18, 2), ls.cntr_value / 1024.0) AS [Total Log File Size on Disk (MB)],
       CONVERT (DECIMAL (18, 2), lu.cntr_value / 1024.0) AS [Log File Used (MB)],
       CAST (CAST (lu.cntr_value AS FLOAT) / CAST (ls.cntr_value AS FLOAT) AS DECIMAL (18, 2)) * 100 AS [Log Used %],
       db.page_verify_option_desc AS [Page Verify Option],
       db.user_access_desc,
       db.state_desc,
       db.containment_desc,
       db.is_auto_create_stats_on,
       db.is_auto_update_stats_on,
       db.is_auto_update_stats_async_on,
       db.is_parameterization_forced,
       db.snapshot_isolation_state_desc,
       db.is_read_committed_snapshot_on,
       db.is_auto_close_on,
       db.is_auto_shrink_on,
       db.target_recovery_time_in_seconds,
       db.is_cdc_enabled,
       db.is_published,
       db.is_distributor,
       db.is_sync_with_backup,
       db.group_database_id,
       db.replica_id,
       db.is_memory_optimized_elevate_to_snapshot_on,
       db.delayed_durability_desc,
       db.is_query_store_on,
       db.is_master_key_encrypted_by_server,
       db.is_encrypted,
       de.encryption_state,
       de.percent_complete,
       de.key_algorithm,
       de.key_length
FROM sys.databases AS db WITH (NOLOCK)
     LEFT OUTER JOIN sys.dm_os_performance_counters AS lu WITH (NOLOCK) ON db.name = lu.instance_name
     LEFT OUTER JOIN sys.dm_os_performance_counters AS ls WITH (NOLOCK) ON db.name = ls.instance_name
     LEFT OUTER JOIN sys.dm_os_performance_counters AS ds WITH (NOLOCK) ON db.name = ds.instance_name
     LEFT OUTER JOIN sys.dm_database_encryption_keys AS de WITH (NOLOCK) ON db.database_id = de.database_id
WHERE lu.counter_name LIKE N'Log File(s) Used Size (KB)%'
      AND ls.counter_name LIKE N'Log File(s) Size (KB)%'
      AND ds.counter_name LIKE N'Data File(s) Size (KB)%'
      AND ls.cntr_value > 0
ORDER BY db.[name]
OPTION (RECOMPILE);
";
            #endregion

            #region m_DatabaseInfo
            // TODO ??
            private const string m_DatabaseInfo = @"
SELECT * FROM sys.databases;
";
            #endregion

            #region m_AlwaysOnStatus
            private const string m_AlwaysOnStatus = @"
SELECT ag.[name] AS [AG Name],
       ar.replica_server_name,
       ar.availability_mode_desc,
       adc.[database_name],
       drs.is_local,
       drs.is_primary_replica,
       drs.synchronization_state_desc,
       drs.is_commit_participant,
       drs.synchronization_health_desc,
       drs.last_commit_time,
       drs.last_sent_time,
       drs.last_received_time,
       drs.last_hardened_time,
       drs.last_redone_time,
       drs.log_send_queue_size,
       drs.log_send_rate,
       drs.redo_queue_size,
       drs.redo_rate
FROM sys.dm_hadr_database_replica_states AS drs WITH (NOLOCK)
     INNER JOIN sys.availability_databases_cluster AS adc WITH (NOLOCK)
		ON drs.group_id = adc.group_id AND drs.group_database_id = adc.group_database_id
     INNER JOIN sys.availability_groups AS ag WITH (NOLOCK)
		ON ag.group_id = drs.group_id
     INNER JOIN sys.availability_replicas AS ar WITH (NOLOCK)
		ON drs.group_id = ar.group_id AND drs.replica_id = ar.replica_id
ORDER BY ag.[name], ar.replica_server_name, adc.[database_name]
OPTION (RECOMPILE);
";
            #endregion

            #region m_spWhoIsActive
            private const string m_spWhoIsActive = @"EXEC sp_WhoIsActive @show_sleeping_spids = 0
	--,@get_plans = 1
	--,@get_full_inner_text = 1
	--,@get_outer_command = 1
	--,@get_locks = 1
	--,@find_block_leaders = 1;";
            #endregion

            #region m_DatabaseBackupDetailedInfo
            private const string m_DatabaseBackupDetailedInfo = @"
WITH LastFullBackups
AS (SELECT s.[database_name] AS DatabaseName,
           MAX(s.backup_set_id) AS BackupSetId
    FROM msdb.dbo.backupset AS s
    WHERE s.[type] = 'D'
    GROUP BY s.[database_name]),
 LastDiffBackups
AS (SELECT s.[database_name] AS DatabaseName,
           MAX(s.backup_set_id) AS BackupSetId
    FROM msdb.dbo.backupset AS s
         INNER JOIN LastFullBackups AS f
             ON s.[database_name] = f.DatabaseName
            AND s.backup_set_id > f.BackupSetId
    WHERE s.[type] = 'I'
    GROUP BY s.[database_name]),
 LastLogBackups
AS (SELECT s.[database_name] AS DatabaseName,
           s.backup_set_id AS BackupSetId
    FROM msdb.dbo.backupset AS s
         INNER JOIN (SELECT DatabaseName,
                 MAX(BackupSetId) AS BackupSetId
          FROM (SELECT DatabaseName,
                       BackupSetId
                FROM LastFullBackups
                UNION ALL
                SELECT DatabaseName,
                       BackupSetId
                FROM LastDiffBackups) AS a
          GROUP BY DatabaseName) AS f
             ON s.[database_name] = f.DatabaseName
            AND s.backup_set_id > f.BackupSetId
    WHERE s.[type] = 'L')
SELECT sd.[name] AS DatabaseName,
       sd.recovery_model_desc,
       bsf.backup_finish_date AS FULL_lastFinishDate,
       CAST (bsf.compressed_backup_size / 1024 / 1024 / 1024. AS NUMERIC (15, 2)) AS FULL_sizeGb,
       bsd.backup_finish_date AS DIFF_lastFinishDate,
       CAST (bsd.compressed_backup_size / 1024 / 1024 / 1024. AS NUMERIC (15, 2)) AS DIFF_sizeGb,
       DATEDIFF(hour, COALESCE (bsd.backup_finish_date, bsf.backup_finish_date), GETDATE()) AS DATA_hoursSinceLastBackup,
       l.backup_finish_date AS LOG_lastFinishDate,
       DATEDIFF(minute, l.backup_finish_date, GETDATE()) AS LOG_minutesSinceLastBackup,
       l.LogBackupCount AS LOG_backupCount,
       CAST (l.compressed_backup_size / 1024 / 1024 / 1024. AS NUMERIC (15, 2)) AS LOG_sizeGb
FROM sys.databases AS sd
     LEFT OUTER JOIN LastFullBackups AS f
     INNER JOIN msdb.dbo.backupset AS bsf
         ON f.BackupSetId = bsf.backup_set_id
         ON f.DatabaseName = sd.[name]
     LEFT OUTER JOIN LastDiffBackups AS d
     INNER JOIN msdb.dbo.backupset AS bsd
         ON d.BackupSetId = bsd.backup_set_id
         ON sd.[name] = d.DatabaseName
     LEFT OUTER JOIN (SELECT l.DatabaseName,
                             COUNT(*) AS LogBackupCount,
                             MAX(bsl.backup_finish_date) AS backup_finish_date,
                             SUM(compressed_backup_size) AS compressed_backup_size
                      FROM LastLogBackups AS l
                           INNER JOIN msdb.dbo.backupset AS bsl
                               ON l.BackupSetId = bsl.backup_set_id
                      GROUP BY l.DatabaseName) AS l
         ON sd.[name] = l.DatabaseName
WHERE sd.[name] <> 'tempdb'
ORDER BY sd.[name];
";
            #endregion

            public string BlockingRequests;
            public string DatabaseLogUsageInfo;
            public string DatabaseInfo;
            public string AlwaysOnStatus;
            public string spWhoIsActive;
            public string DatabaseBackupDetailedInfo;

            public HealthDashboardServerQueryTexts()
            {
                BlockingRequests = m_BlockingRequestsQuery;
                DatabaseLogUsageInfo = m_DatabaseLogUsageInfo;
                DatabaseInfo = m_DatabaseInfo;
                AlwaysOnStatus = m_AlwaysOnStatus;
                spWhoIsActive = m_spWhoIsActive;
                DatabaseBackupDetailedInfo = m_DatabaseBackupDetailedInfo;
            }
        }

        public class SmtpSettings
        {
            public bool hasBeenConfiguredAndTested;
            public string Username;
            public string Password;
            public string ServerName;
            public int Port;
            public bool EnableSsl;
        }

        public class ExcelExportSettings
        {
            public bool includeSourceQuery = false;
            public bool addAutofilter = false;
            public bool exportBoolsAsNumbers = false;
            public string defaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            public string defaultFileName = "DataExport_{yyyyMMdd_HHmmss}.xlsx";

            public string GetDefaultDirectory()
            {
                if (!string.IsNullOrEmpty(defaultDirectory)
                    && Directory.Exists(defaultDirectory))
                {
                    return defaultDirectory;
                }

                return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }

            public string GetDefaultFileName()
            {
                if (!string.IsNullOrWhiteSpace(defaultFileName))
                {
                    return defaultFileName;
                }

                return "DataExport_{yyyyMMdd_HHmmss}.xlsx";
            }
        }

        public class TSqlCodeFormatSettings
        {
            public bool preserveComments = false;
            public bool removeNewLineAfterJoin = false;
            public bool addTabAfterJoinOn = false;
            public bool moveCrossJoinToNewLine = false;
            public bool formatCaseAsMultiline = false;
            public bool addNewLineBetweenStatementsInBlocks = false;
            public bool breakSprocParametersPerLine = false;
            public bool uppercaseBuiltInFunctions = false;
            public bool unindentBeginEndBlocks = false;
            public bool breakVariableDefinitionsPerLine = false;
            public bool breakSprocDefinitionParametersPerLine = false;
            public bool breakSelectFieldsAfterTopAndUnindent = false;

            public bool HasAnyFormattingEnabled()
            {
                // preserveComments option is not included because it belongs to a separate code branch.
                return removeNewLineAfterJoin
                    || addTabAfterJoinOn
                    || moveCrossJoinToNewLine
                    || formatCaseAsMultiline
                    || addNewLineBetweenStatementsInBlocks
                    || breakSprocParametersPerLine
                    || uppercaseBuiltInFunctions
                    || unindentBeginEndBlocks
                    || breakVariableDefinitionsPerLine
                    || breakSprocDefinitionParametersPerLine
                    || breakSelectFieldsAfterTopAndUnindent;
            }
        }

        public class FrequentlyUsedEmail
        {
            public string EmailAddress;
            public int UsedTimes;

            public FrequentlyUsedEmail(string emailAddress, int usedTimes)
            {
                EmailAddress = emailAddress;
                UsedTimes = usedTimes;
            }
        }

        public static byte[] Protect(byte[] data)
        {
            try
            {
                // Use the current user scope to encrypt the data.
                return ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
            }
            catch (CryptographicException e)
            {
                Console.WriteLine($"A cryptographic error occurred: {e.Message}");
                return null;
            }
        }

        public static byte[] Unprotect(byte[] data)
        {
            try
            {
                // Decrypt the data using the current user scope.
                return ProtectedData.Unprotect(data, null, DataProtectionScope.CurrentUser);
            }
            catch (CryptographicException e)
            {
                Console.WriteLine($"A cryptographic error occurred: {e.Message}");
                return null;
            }
        }

        private static RegistryKey GetRoot()
        {
            var settingKeyRoot = Registry.CurrentUser.CreateSubKey(@"AxialSqlTools");
            var settingsKey = settingKeyRoot.CreateSubKey("Settings");

            return settingsKey;
        }

        private static string GetRegisterValue(string parameter)
        {
            try
            {
                using (var rootKey = GetRoot())
                {
                    var value = rootKey.GetValue(parameter);
                    return value?.ToString() ?? string.Empty;
                }
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        private static bool SaveRegisterValue(string parameterName, string parameterValue)
        {
            try
            {
                using (var rootKey = GetRoot())
                {
                    rootKey.SetValue(parameterName, parameterValue);
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        // ----------------------------------------------------------------------
        public static string GetMyEmail()
        {
            return GetRegisterValue("MyEmail");
        }
        public static bool SaveMyEmail(string myEmail)
        {
            return SaveRegisterValue("MyEmail", myEmail);
        }
        // ----------------------------------------------------------------------
        public static SmtpSettings GetSmtpSettings()
        {

            string password = "";
            try
            {
                string encPassword = GetRegisterValue("SMTP_Password");
                byte[] decryptedData = Unprotect(Convert.FromBase64String(encPassword));
                password = Encoding.UTF8.GetString(decryptedData);
            } catch {
            }

            SmtpSettings smtpSettings = new SmtpSettings();            
            smtpSettings.hasBeenConfiguredAndTested = true; //TODO
            smtpSettings.Username = GetRegisterValue("SMTP_Username");
            smtpSettings.Password = password;
            smtpSettings.ServerName = GetRegisterValue("SMTP_Server");

            smtpSettings.Port = 587;
            int savedPort;
            bool success = int.TryParse(GetRegisterValue("SMTP_Port"), out savedPort);
            if (success) smtpSettings.Port = savedPort;

            smtpSettings.EnableSsl = true;
            bool enableSsl;
            success = bool.TryParse(GetRegisterValue("SMTP_EnableSSL"), out enableSsl);
            if (success) smtpSettings.EnableSsl = enableSsl;

            return smtpSettings;
        }

        public static bool SaveSmtpSettings(SmtpSettings smtpSettings)
        {

            byte[] encPassword = Protect(Encoding.UTF8.GetBytes(smtpSettings.Password));

            SaveRegisterValue("SMTP_Username", smtpSettings.Username);
            SaveRegisterValue("SMTP_Password", Convert.ToBase64String(encPassword));
            SaveRegisterValue("SMTP_Server", smtpSettings.ServerName);
            SaveRegisterValue("SMTP_Port", smtpSettings.Port.ToString());
            SaveRegisterValue("SMTP_EnableSSL", smtpSettings.EnableSsl.ToString());

            return true;

        }

        public static string GetTemplatesFolder()
        {
            var folder = GetRegisterValue("ScriptTemplatesFolder");

            if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder))
            {
                folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "AxialSqlToolsTemplates");

                SaveTemplatesFolder(folder);

                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }
            }

            return folder;
        }

        public static bool SaveTemplatesFolder(string folder)
        {
            return SaveRegisterValue("ScriptTemplatesFolder", folder);
        }

        public static List<FrequentlyUsedEmail> GetFrequentlyUsedEmails()
        {

            try
            {

                string JsonString = GetRegisterValue("FrequentlyUsedEmails");
                var deserializedEmailList = JsonConvert.DeserializeObject<List<FrequentlyUsedEmail>>(JsonString);

                if (deserializedEmailList == null)
                    return new List<FrequentlyUsedEmail>();

                return deserializedEmailList;

            } catch {}

            return new List<FrequentlyUsedEmail>();

        }

        public static void SaveFrequentlyUsedEmails(List<string> UsedEmails)
        {

            List<FrequentlyUsedEmail> existingEmails = GetFrequentlyUsedEmails();

            foreach (var email in UsedEmails)
            {
                var existingEmail = existingEmails.FirstOrDefault(e => e.EmailAddress == email);
                if (existingEmail != null)
                {
                    existingEmail.UsedTimes++;
                }
                else
                {
                    existingEmails.Add(new FrequentlyUsedEmail(email, 1));
                }
            }

            existingEmails.Sort((email1, email2) => email2.UsedTimes.CompareTo(email1.UsedTimes));

            // Keep only the top 30 most frequently used emails
            if (existingEmails.Count > 30)
                existingEmails = existingEmails.Take(30).ToList();

            string json = JsonConvert.SerializeObject(existingEmails);

            SaveRegisterValue("FrequentlyUsedEmails", json);

        }

        public static string GetOpenAiApiKey()
        {
            string key = "";
            try
            {
                string encKey = GetRegisterValue("OpenAI_ApiKeyEnc");
                byte[] decryptedData = Unprotect(Convert.FromBase64String(encKey));
                key = Encoding.UTF8.GetString(decryptedData);
            }
            catch
            {
            }

            return key;
        }
        public static bool SaveOpenAiApiKey(string ApiKey)
        {
            byte[] encKey = Protect(Encoding.UTF8.GetBytes(ApiKey));
            return SaveRegisterValue("OpenAI_ApiKeyEnc", Convert.ToBase64String(encKey));
        }


        public static bool GetUseSnippets()
        {
            bool result = false;
            bool success = bool.TryParse(GetRegisterValue("UseSnippets"), out result);
            return result;
        }
        public static string GetSnippetFolder()
        {
            var folder = GetRegisterValue("SnippetFolder");
            /*
            if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder))
            {
                folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "AxialSqlToolsTemplates");

                SaveTemplatesFolder(folder);

                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }
            }
            */
            return folder;
        }
        public static bool SaveSnippetUse(bool UseSnippets, string SnippetFolder)
        {
            return SaveUseSnippets(UseSnippets) && SaveSnippetFolder(SnippetFolder);
        }
        public static bool SaveUseSnippets(bool UseSnippets)
        {
            return SaveRegisterValue("UseSnippets", UseSnippets.ToString());
        }
        public static bool SaveSnippetFolder(string folder)
        {
            return SaveRegisterValue("SnippetFolder", folder);
        }


        public static string GetQueryHistoryConnectionString()
        {
            string key = "";
            try
            {
                string encKey = GetRegisterValue("QueryHistoryConnectionString");
                byte[] decryptedData = Unprotect(Convert.FromBase64String(encKey));
                key = Encoding.UTF8.GetString(decryptedData);
            }
            catch
            {
            }

            return key;
        }
        public static bool SaveQueryHistoryConnectionString(string connectionString)
        {
            byte[] encKey = Protect(Encoding.UTF8.GetBytes(connectionString));
            return SaveRegisterValue("QueryHistoryConnectionString", Convert.ToBase64String(encKey));
        }

        public static string GetQueryHistoryTableNameOrDefault()
        {
            string qhTable = GetRegisterValue("QueryHistoryTableName");

            if (string.IsNullOrEmpty(qhTable))
            {
                qhTable = "QueryHistory";
            }
            return qhTable;
        }

        public static string GetQueryHistoryTableName()
        {
            return GetRegisterValue("QueryHistoryTableName");
        }
        public static bool SaveQueryHistoryTableName(string qhTableName)
        {
            return SaveRegisterValue("QueryHistoryTableName", qhTableName);
        }

        public static ExcelExportSettings GetExcelExportSettings()
        {
            try
            {
                // read the JSON blob (or empty string)
                string json = GetRegisterValue("ExcelExportSettings");
                if (string.IsNullOrEmpty(json))
                    return new ExcelExportSettings();

                // deserialize; on failure fall back to defaults
                var settings = JsonConvert.DeserializeObject<ExcelExportSettings>(json);
                return settings ?? new ExcelExportSettings();
            }
            catch
            {
                return new ExcelExportSettings();
            }
        }

        public static bool SaveExcelExportSettings(ExcelExportSettings settings)
        {
            try
            {
                // serialize to JSON and write to registry
                string json = JsonConvert.SerializeObject(settings);
                return SaveRegisterValue("ExcelExportSettings", json);
            }
            catch
            {
                return false;
            }
        }

        public static TSqlCodeFormatSettings GetTSqlCodeFormatSettings()
        {
            try
            {
                string json = GetRegisterValue("TSqlCodeFormatSettings");
                if (string.IsNullOrEmpty(json))
                    return new TSqlCodeFormatSettings();

                var settings = JsonConvert.DeserializeObject<TSqlCodeFormatSettings>(json);
                return settings ?? new TSqlCodeFormatSettings();
            }
            catch
            {
                return new TSqlCodeFormatSettings();
            }
        }

        public static bool SaveTSqlCodeFormatSettings(TSqlCodeFormatSettings settings)
        {
            try
            {
                string json = JsonConvert.SerializeObject(settings);
                return SaveRegisterValue("TSqlCodeFormatSettings", json);
            }
            catch
            {
                return false;
            }
        }


    }
}
