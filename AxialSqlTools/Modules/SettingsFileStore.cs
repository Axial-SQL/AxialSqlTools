using NLog;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Security.Cryptography;

namespace AxialSqlTools
{
    internal static class SettingsFileStore
    {
        private static readonly string PathToFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AxialSqlTools", "settings.json");

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static readonly object CacheLock = new object();
        private static readonly Timer RefreshTimer = new Timer(RefreshCache, null, Timeout.Infinite, Timeout.Infinite);
        private static Dictionary<string, string> _snapshot;
        private static string _lastLoadedJson;
        private static bool _readFailureReported;

        [ThreadStatic]
        private static string _lastSaveError;
        internal static string LastSaveError => _lastSaveError;

        public static string GetValue(string name)
        {
            var snapshot = Volatile.Read(ref _snapshot);
            if (snapshot == null)
            {
                lock (CacheLock)
                {
                    if (_snapshot == null)
                    {
                        RefreshCacheCore();
                        if (_snapshot == null)
                            Publish(new JObject());
                        RefreshTimer.Change(2000, 2000);
                    }
                    snapshot = _snapshot;
                }
            }

            return snapshot.TryGetValue(name, out string value) ? value : string.Empty;
        }

        private static void RefreshCache(object state)
        {
            // Poll in the background, including when the directory does not exist yet.
            // Never queue overlapping refreshes when a redirected folder is slow.
            if (!Monitor.TryEnter(CacheLock))
                return;
            try
            {
                RefreshCacheCore();
            }
            finally
            {
                Monitor.Exit(CacheLock);
            }
        }

        private static void RefreshCacheCore()
        {
            try
            {
                string json = ReadText();
                if (!string.Equals(json, _lastLoadedJson, StringComparison.Ordinal))
                {
                    Publish(JObject.Parse(json));
                    _lastLoadedJson = json;
                }
                _readFailureReported = false;
            }
            catch (Exception ex)
            {
                // Keep the last good snapshot while an external edit is incomplete.
                if (!_readFailureReported)
                    Logger.Warn("Could not reload settings file {0} ({1}). Keeping the last valid settings.",
                        PathToFile, ex.GetType().Name);
                _readFailureReported = true;
            }
        }

        private static void Publish(JObject settings)
        {
            var snapshot = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var property in settings.Properties())
            {
                var value = property.Value;
                snapshot[property.Name] = value.Type == JTokenType.Null ? string.Empty
                    : value.Type == JTokenType.String ? value.Value<string>()
                    : value.ToString(Formatting.None);
            }
            // Published dictionaries are never mutated; readers need no lock or disk access.
            Volatile.Write(ref _snapshot, snapshot);
        }

        internal static bool ReportSaveFailure(Exception ex)
        {
            _lastSaveError = ex is UnauthorizedAccessException ? "Access to the settings file was denied. Check its permissions."
                : ex is JsonException ? "The settings file contains invalid JSON. Correct it before saving."
                : ex is CryptographicException ? "The credentials could not be encrypted for the current Windows user."
                : ex is TimeoutException ? "Another SSMS process is saving settings. Please try again."
                : ex is IOException ? "The settings file could not be written. Check the folder, available space, and file access."
                : "The settings could not be saved. See the Axial SQL Tools log for details.";
            // Do not log JSON contents, tokens, passwords, or exception messages containing them.
            Logger.Error("Settings save failed ({0}): {1} File: {2}. Stack: {3}",
                ex.GetType().Name, _lastSaveError, PathToFile, ex.StackTrace);
            return false;
        }

        public static bool SaveValue<T>(string name, T value)
        {
            return SaveValues(new Dictionary<string, object> { [name] = value });
        }

        public static bool SaveValues(IDictionary<string, object> values)
        {
            _lastSaveError = null;
            try
            {
                // Serialize read/modify/write operations across SSMS processes.
                using (var mutex = new Mutex(false, @"Local\AxialSQL.SettingsFileStore"))
                {
                    bool acquired = false;
                    try
                    {
                        try
                        {
                            acquired = mutex.WaitOne(TimeSpan.FromSeconds(5));
                        }
                        catch (AbandonedMutexException)
                        {
                            acquired = true;
                        }

                        if (!acquired)
                            return ReportSaveFailure(new TimeoutException());

                        lock (CacheLock)
                        {
                            // Always reload for writes; never overwrite a malformed file with cached data.
                            var settings = JObject.Parse(ReadText());
                            foreach (var pair in values)
                                settings[pair.Key] = pair.Value == null
                                    ? JValue.CreateNull()
                                    : JToken.FromObject(pair.Value);

                            Write(settings);
                            Publish(settings);
                            _lastLoadedJson = null;
                            RefreshTimer.Change(2000, 2000);
                        }
                        return true;
                    }
                    finally
                    {
                        if (acquired)
                            mutex.ReleaseMutex();
                    }
                }
            }
            catch (Exception ex)
            {
                return ReportSaveFailure(ex);
            }
        }

        private static string ReadText()
        {
            try
            {
                // Readers can finish reading the old snapshot during an atomic replace.
                using (var stream = new FileStream(PathToFile, FileMode.Open, FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete))
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
            catch (FileNotFoundException)
            {
                return "{}";
            }
            catch (DirectoryNotFoundException)
            {
                return "{}";
            }
        }

        private static void Write(JObject settings)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(PathToFile));
            string temporaryPath = PathToFile + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllText(temporaryPath, settings.ToString(Formatting.Indented));
                if (File.Exists(PathToFile))
                    File.Replace(temporaryPath, PathToFile, null);
                else
                    File.Move(temporaryPath, PathToFile);
            }
            finally
            {
                if (File.Exists(temporaryPath))
                    File.Delete(temporaryPath);
            }
        }
    }
}
