using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace AxialSqlTools
{
    internal static class SettingsFileStore
    {
        private static readonly string PathToFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AxialSQL", "settings.json");

        public static string GetValue(string name)
        {
            try
            {
                var value = Load()[name];
                if (value == null || value.Type == JTokenType.Null)
                    return string.Empty;

                return value.Type == JTokenType.String
                    ? value.Value<string>()
                    : value.ToString(Formatting.None);
            }
            catch
            {
                // Keep the existing defaults when the file cannot be read.
                return string.Empty;
            }
        }

        public static bool SaveValue<T>(string name, T value)
        {
            return SaveValues(new Dictionary<string, object> { [name] = value });
        }

        public static bool SaveValues(IDictionary<string, object> values)
        {
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
                            return false;

                        // Reload to preserve settings changed by another SSMS process.
                        // A malformed or unreadable file must not be overwritten.
                        var settings = Load();
                        foreach (var pair in values)
                            settings[pair.Key] = pair.Value == null
                                ? JValue.CreateNull()
                                : JToken.FromObject(pair.Value);

                        Write(settings);
                        return true;
                    }
                    finally
                    {
                        if (acquired)
                            mutex.ReleaseMutex();
                    }
                }
            }
            catch
            {
                return false;
            }
        }

        private static JObject Load()
        {
            try
            {
                // Readers can finish reading the old snapshot during an atomic replace.
                using (var stream = new FileStream(PathToFile, FileMode.Open, FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete))
                using (var reader = new StreamReader(stream))
                    return JObject.Parse(reader.ReadToEnd());
            }
            catch (FileNotFoundException)
            {
                return new JObject();
            }
            catch (DirectoryNotFoundException)
            {
                return new JObject();
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
