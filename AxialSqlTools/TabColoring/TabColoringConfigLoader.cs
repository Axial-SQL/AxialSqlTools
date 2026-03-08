using System;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization.Json;
using System.Text;

namespace AxialSqlTools.TabColoring
{
    internal static class TabColoringConfigLoader
    {
        private const string DefaultConfigResourceName = "AxialSqlTools.TabColoring.DefaultTabColoringConfig.json";

        public static string GetUserConfigPath()
        {
            string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(docs, "AxialSqlTools", "TabColoringConfig.json");
        }

        public static void EnsureDefaultConfigExists()
        {
            try
            {
                string path = GetUserConfigPath();
                string dir = Path.GetDirectoryName(path);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                if (File.Exists(path))
                {
                    return;
                }

                string defaultJson = ReadDefaultConfigJsonOrNull();
                if (string.IsNullOrWhiteSpace(defaultJson))
                {
                    defaultJson = "{\n  \"settings\": { \"enabled\": false },\n  \"queryRules\": []\n}\n";
                }

                File.WriteAllText(path, defaultJson, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            }
            catch
            {
            }
        }

        public static TabColoringConfig LoadOrNull()
        {
            try
            {
                string path = GetUserConfigPath();
                if (!File.Exists(path))
                {
                    return null;
                }

                string jsonText = File.ReadAllText(path, Encoding.UTF8);
                if (string.IsNullOrWhiteSpace(jsonText))
                {
                    return null;
                }

                jsonText = jsonText.TrimStart('\uFEFF', '\u200B', '\u0000', ' ', '\t', '\r', '\n');
                byte[] utf8Bytes = Encoding.UTF8.GetBytes(jsonText);
                using (var stream = new MemoryStream(utf8Bytes))
                {
                    var serializer = new DataContractJsonSerializer(typeof(TabColoringConfig), new DataContractJsonSerializerSettings
                    {
                        UseSimpleDictionaryFormat = true
                    });

                    return serializer.ReadObject(stream) as TabColoringConfig;
                }
            }
            catch
            {
                return null;
            }
        }

        private static string ReadDefaultConfigJsonOrNull()
        {
            try
            {
                Assembly asm = typeof(TabColoringConfigLoader).Assembly;
                using (Stream s = asm.GetManifestResourceStream(DefaultConfigResourceName))
                {
                    if (s == null)
                    {
                        return null;
                    }

                    using (var reader = new StreamReader(s, Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            catch
            {
                return null;
            }
        }
    }
}
