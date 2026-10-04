using System.Collections.Generic;

namespace AxialSqlTools
{
    // Keep persistence tests isolated from the user's actual extension settings.
    internal static class SettingsFileStore
    {
        internal static readonly Dictionary<string, string> Values = new Dictionary<string, string>();
        internal static bool FailSave;
        internal static string LastSaveError => FailSave ? "Simulated write failure" : null;
        internal static string GetValue(string key)
        {
            string value;
            return Values.TryGetValue(key, out value) ? value : null;
        }
        internal static bool SaveValue<T>(string key, T value)
        {
            if (FailSave) return false;
            Values[key] = value == null ? null : value.ToString();
            return true;
        }
        internal static void Reset() { Values.Clear(); FailSave = false; }
    }
}
