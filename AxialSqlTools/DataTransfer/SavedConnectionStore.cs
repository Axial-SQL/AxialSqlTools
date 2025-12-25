using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace AxialSqlTools
{
    public static class SavedConnectionStore
    {
        private static readonly string PathToFile =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AxialSQL", "data-transfer-connections.json");

        public static List<SettingsManager.DataTransferSavedConnection> Load()
        {
            try
            {
                if (!File.Exists(PathToFile))
                {
                    return new List<SettingsManager.DataTransferSavedConnection>();
                }

                var json = File.ReadAllText(PathToFile);
                return JsonConvert.DeserializeObject<List<SettingsManager.DataTransferSavedConnection>>(json)
                       ?? new List<SettingsManager.DataTransferSavedConnection>();
            }
            catch
            {
                return new List<SettingsManager.DataTransferSavedConnection>();
            }
        }

        public static void Save(IEnumerable<SettingsManager.DataTransferSavedConnection> connections)
        {
            var directory = Path.GetDirectoryName(PathToFile);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(PathToFile, JsonConvert.SerializeObject(connections, Formatting.Indented));
        }
    }
}
