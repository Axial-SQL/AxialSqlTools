using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using System.Threading.Tasks;

namespace AxialSqlTools
{
    class SettingsManager
    {

        public class HealthDashboardServerQueryTexts
        {
            #region m_BlockingRequestsQuery
            private const string m_BlockingRequestsQuery = @"
WITH Blocking
AS (SELECT [session_Id] AS [Spid],
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
           Hostname,
           cpu_time,
           reads,
           start_time,
           sp.login_time,
           sp.last_batch,
           sp.cmd
    FROM sys.dm_exec_requests AS er
         INNER JOIN sys.sysprocesses AS sp ON er.[session_id] = sp.spid 
         CROSS APPLY sys.dm_exec_sql_text(er.[sql_handle]) AS qt
    WHERE [session_Id] > 50 AND [session_Id] NOT IN (@@SPID))
SELECT * FROM blocking WHERE blocked <> 0
UNION ALL
SELECT * FROM blocking WHERE spid IN (SELECT blocked FROM blocking WHERE blocked <> 0);
";
            #endregion

            public string BlockingRequests;

            public HealthDashboardServerQueryTexts()
            {
                BlockingRequests = m_BlockingRequestsQuery;
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

        private static byte[] Protect(byte[] data)
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

        private static byte[] Unprotect(byte[] data)
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

    }
}
