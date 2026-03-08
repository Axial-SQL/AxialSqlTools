using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace SSMS_EnvTabs
{
    internal static class SsmsSettingsUpdater
    {
        private const string ColorizationKey = "environment.tabs.documentTabs.colorization";

        public static void EnsureRegexTabColorizationEnabled(bool enableAutoColor)
        {
            if (!enableAutoColor)
            {
                return;
            }

            try
            {
                string baseDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "SSMS");
                if (!Directory.Exists(baseDir))
                {
                    return;
                }

                var settingsFiles = Directory.GetDirectories(baseDir)
                    .Select(dir => Path.Combine(dir, "settings.json"))
                    .Where(File.Exists)
                    .ToList();

                foreach (var settingsPath in settingsFiles)
                {
                    TryUpdateSettingsFile(settingsPath);
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"SsmsSettingsUpdater: failed to update settings.json: {ex.Message}");
            }
        }

        private static void TryUpdateSettingsFile(string settingsPath)
        {
            try
            {
                string text = File.ReadAllText(settingsPath);
                if (string.IsNullOrWhiteSpace(text))
                {
                    return;
                }

                string pattern = $"(\"{Regex.Escape(ColorizationKey)}\"\\s*:\\s*\")([^\"]*)(\")";
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    string currentValue = match.Groups[2].Value;
                    if (string.Equals(currentValue, "regex", StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }

                    string replaced = Regex.Replace(text, pattern, m => $"\"{ColorizationKey}\":\"regex\"", RegexOptions.IgnoreCase);
                    File.WriteAllText(settingsPath, replaced);
                    EnvTabsLog.Info($"SsmsSettingsUpdater: Updated colorization setting in '{settingsPath}'.");
                    return;
                }

                // Not present: insert before final closing brace
                int lastBrace = text.LastIndexOf('}');
                if (lastBrace < 0)
                {
                    return;
                }

                string newline = text.Contains("\r\n") ? "\r\n" : "\n";
                string prefix = text.Substring(0, lastBrace);

                int i = prefix.Length - 1;
                while (i >= 0 && char.IsWhiteSpace(prefix[i])) i--;
                char lastChar = i >= 0 ? prefix[i] : '{';
                bool addComma = lastChar != '{' && lastChar != ',';

                string insertion = (addComma ? "," : string.Empty) + newline + "  \"" + ColorizationKey + "\": \"regex\"" + newline;
                string updated = prefix + insertion + text.Substring(lastBrace);
                File.WriteAllText(settingsPath, updated);
                EnvTabsLog.Info($"SsmsSettingsUpdater: Added colorization setting in '{settingsPath}'.");
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"SsmsSettingsUpdater: failed to update '{settingsPath}': {ex.Message}");
            }
        }
    }
}