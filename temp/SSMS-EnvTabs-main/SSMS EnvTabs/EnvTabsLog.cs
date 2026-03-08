using Microsoft.VisualStudio.Shell;
using System;
using System.IO;
using System.Text;

namespace SSMS_EnvTabs
{
    internal static class EnvTabsLog
    {
        private static readonly string LogPath = ResolveLogPath();
        private static readonly object LogLock = new object();
        public static bool Enabled { get; set; } = false;
        public static bool VerboseEnabled { get; set; } = false;

        private static string ResolveLogPath()
        {
            try
            {
                string baseDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                if (string.IsNullOrWhiteSpace(baseDir))
                {
                    baseDir = Path.GetTempPath();
                }

                return Path.Combine(baseDir, "SSMS EnvTabs", "runtime.log");
            }
            catch
            {
                return Path.Combine(Path.GetTempPath(), "SSMS EnvTabs", "runtime.log");
            }
        }

        public static void Info(string message)
        {
            if (!Enabled) return;
            try
            {
                if (ThreadHelper.CheckAccess())
                {
                    ActivityLog.LogInformation("SSMS EnvTabs", message ?? string.Empty);
                }
            }
            catch
            {
                // Best-effort logging only.
            }

            WriteToFile("INFO", message);
        }

        public static void Error(string message)
        {
            if (!Enabled) return;
            try
            {
                if (ThreadHelper.CheckAccess())
                {
                    ActivityLog.LogError("SSMS EnvTabs", message ?? string.Empty);
                }
            }
            catch
            {
                // Best-effort logging only.
            }

            WriteToFile("ERROR", message);
        }

        public static void Verbose(string message)
        {
            if (!Enabled || !VerboseEnabled) return;
            Info(message);
        }

        private static void WriteToFile(string level, string message)
        {
            try
            {
                string dir = Path.GetDirectoryName(LogPath);
                if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{level}] {message}";
                lock (LogLock)
                {
                    File.AppendAllText(LogPath, line + Environment.NewLine, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                }
            }
            catch
            {
                // Best-effort logging only.
            }
        }
    }
}
