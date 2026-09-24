using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;

namespace AxialSqlTools
{
    internal sealed class SsmsFormatterSettingValue
    {
        public string Name { get; }
        public string Value { get; }

        internal SsmsFormatterSettingValue(string name, string value)
        {
            Name = name;
            Value = value;
        }
    }

    internal sealed class SsmsFormatterSettingsSnapshot
    {
        internal const string Scope = "Global SSMS options. Document .editorconfig overrides are not included.";
        internal string Versions { get; private set; }
        internal IReadOnlyList<SsmsFormatterSettingValue> Settings { get; private set; }

        internal static SsmsFormatterSettingsSnapshot Capture(object settings, Assembly formatterAssembly)
        {
            if (settings == null) throw new ArgumentNullException(nameof(settings));
            var values = new List<SsmsFormatterSettingValue>();
            ReadMembers(settings, string.Empty, values, 0);
            if (values.Count == 0)
                throw new InvalidOperationException("This SSMS formatter version exposes no readable settings.");

            return new SsmsFormatterSettingsSnapshot
            {
                Versions = "SSMS " + GetSsmsVersion() + " | Formatter " + GetVersion(formatterAssembly)
                    + " | Axial SQL Tools " + GetVersion(typeof(SsmsFormatterSettingsSnapshot).Assembly),
                Settings = values.OrderBy(value => value.Name, StringComparer.Ordinal).ToArray()
            };
        }

        private static void ReadMembers(object source, string prefix, List<SsmsFormatterSettingValue> values, int depth)
        {
            const BindingFlags flags = BindingFlags.Public | BindingFlags.Instance;
            // Inspect the loaded settings object, not a hard-coded list or the stale singleton.
            // Newly added SSMS options therefore appear automatically.
            foreach (var property in source.GetType().GetProperties(flags)
                .Where(property => property.GetGetMethod() != null && property.GetIndexParameters().Length == 0))
                ReadValue(prefix + property.Name, () => property.GetValue(source), values, depth);
            foreach (var field in source.GetType().GetFields(flags))
                ReadValue(prefix + field.Name, () => field.GetValue(source), values, depth);
        }

        private static void ReadValue(string name, Func<object> read, List<SsmsFormatterSettingValue> values, int depth)
        {
            try
            {
                var value = read();
                if (value == null)
                    values.Add(new SsmsFormatterSettingValue(name, "(not set)"));
                else if (value is string text)
                    values.Add(new SsmsFormatterSettingValue(name, text.Length == 0 ? "(empty)" :
                        text.Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t")));
                else if (value.GetType().IsEnum || value is IConvertible || value is Guid)
                    values.Add(new SsmsFormatterSettingValue(name, Convert.ToString(value, CultureInfo.InvariantCulture)));
                else
                {
                    int count = values.Count;
                    if (depth < 3) ReadMembers(value, name + ".", values, depth + 1);
                    if (values.Count == count)
                        values.Add(new SsmsFormatterSettingValue(name, "(unavailable: " + value.GetType().Name + ")"));
                }
            }
            catch (Exception ex)
            {
                // Keep the other settings visible and never label a failed read as a default.
                var cause = (ex as TargetInvocationException)?.InnerException ?? ex;
                values.Add(new SsmsFormatterSettingValue(name, "(unavailable: " + cause.GetType().Name + ")"));
            }
        }

        private static string GetSsmsVersion()
        {
            try
            {
                using (var process = Process.GetCurrentProcess())
                    return process.MainModule.FileVersionInfo.ProductVersion ?? "unknown";
            }
            catch { return "unknown"; }
        }

        private static string GetVersion(Assembly assembly)
        {
            try
            {
                return FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion
                    ?? assembly.GetName().Version?.ToString() ?? "unknown";
            }
            catch { return assembly.GetName().Version?.ToString() ?? "unknown"; }
        }

        internal string ToText(string mode)
        {
            var text = new StringBuilder();
            text.AppendLine("SSMS formatter settings").AppendLine(Versions).AppendLine(Scope).AppendLine(mode).AppendLine();
            foreach (var setting in Settings)
                text.Append(setting.Name).Append(" = ").AppendLine(setting.Value);
            return text.ToString();
        }
    }
}
