using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace AxialSqlTools
{
    public static class SnippetService
    {
        private static readonly object _lock = new object();
        private static List<SnippetItem> _snippets = new List<SnippetItem>();
        private static Dictionary<string, SnippetItem> _snippetDictionary = new Dictionary<string, SnippetItem>(StringComparer.OrdinalIgnoreCase);
        private static bool _loaded = false;

        public static string SnippetFilePath
        {
            get
            {
                string folder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "AxialSqlTools");
                return Path.Combine(folder, "snippets.json");
            }
        }

        public static Dictionary<string, SnippetItem> SnippetDictionary
        {
            get
            {
                EnsureLoaded();
                lock (_lock)
                {
                    return new Dictionary<string, SnippetItem>(_snippetDictionary, StringComparer.OrdinalIgnoreCase);
                }
            }
        }

        public static List<SnippetItem> GetAllSnippets()
        {
            EnsureLoaded();
            lock (_lock)
            {
                return new List<SnippetItem>(_snippets);
            }
        }

        public static void ReloadSnippets()
        {
            lock (_lock)
            {
                _snippets = LoadFromFile();
                RebuildDictionary();
                _loaded = true;
            }
        }

        public static void AddSnippet(SnippetItem snippet)
        {
            EnsureLoaded();
            lock (_lock)
            {
                _snippets.Add(snippet);
                SaveToFile(_snippets);
                RebuildDictionary();
            }
        }

        public static void UpdateSnippet(SnippetItem snippet)
        {
            EnsureLoaded();
            lock (_lock)
            {
                int index = _snippets.FindIndex(s => s.Id == snippet.Id);
                if (index >= 0)
                {
                    _snippets[index] = snippet;
                    SaveToFile(_snippets);
                    RebuildDictionary();
                }
            }
        }

        public static void DeleteSnippet(string id)
        {
            EnsureLoaded();
            lock (_lock)
            {
                _snippets.RemoveAll(s => s.Id == id);
                SaveToFile(_snippets);
                RebuildDictionary();
            }
        }

        public static void ImportFromLegacyFolder(string folderPath)
        {
            if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
                return;

            EnsureLoaded();
            lock (_lock)
            {
                var existingPrefixes = new HashSet<string>(_snippets.Select(s => s.Prefix), StringComparer.OrdinalIgnoreCase);
                var files = Directory.EnumerateFiles(folderPath, "*.sql");

                foreach (var file in files)
                {
                    var fi = new FileInfo(file);
                    if (fi.Length >= 1024 * 1024)
                        continue;

                    string prefix = Path.GetFileNameWithoutExtension(fi.Name);
                    if (existingPrefixes.Contains(prefix))
                        continue;

                    string body = File.ReadAllText(fi.FullName);
                    _snippets.Add(new SnippetItem(prefix, prefix, body));
                    existingPrefixes.Add(prefix);
                }

                SaveToFile(_snippets);
                RebuildDictionary();
            }
        }

        public static void SaveAllSnippets(List<SnippetItem> snippets)
        {
            lock (_lock)
            {
                _snippets = new List<SnippetItem>(snippets);
                SaveToFile(_snippets);
                RebuildDictionary();
            }
        }

        private static void EnsureLoaded()
        {
            if (!_loaded)
            {
                ReloadSnippets();
            }
        }

        private static void RebuildDictionary()
        {
            _snippetDictionary = new Dictionary<string, SnippetItem>(StringComparer.OrdinalIgnoreCase);
            foreach (var snippet in _snippets)
            {
                if (!string.IsNullOrEmpty(snippet.Prefix) && !_snippetDictionary.ContainsKey(snippet.Prefix))
                {
                    _snippetDictionary[snippet.Prefix] = snippet;
                }
            }
        }

        private static List<SnippetItem> LoadFromFile()
        {
            try
            {
                string filePath = SnippetFilePath;
                if (File.Exists(filePath))
                {
                    string json = File.ReadAllText(filePath);
                    var snippets = JsonConvert.DeserializeObject<List<SnippetItem>>(json);
                    if (snippets != null)
                        return snippets;
                }
            }
            catch
            {
            }

            return GetDefaultSnippets();
        }

        private static void SaveToFile(List<SnippetItem> snippets)
        {
            try
            {
                string filePath = SnippetFilePath;
                string folder = Path.GetDirectoryName(filePath);
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);

                string json = JsonConvert.SerializeObject(snippets, Formatting.Indented);
                File.WriteAllText(filePath, json);
            }
            catch
            {
            }
        }

        private static List<SnippetItem> GetDefaultSnippets()
        {
            return new List<SnippetItem>
            {
                new SnippetItem("sel", "Select all from table", "SELECT *\r\nFROM #\r\nWHERE 1=1"),
                new SnippetItem("selc", "Select count", "SELECT COUNT(*)\r\nFROM #"),
                new SnippetItem("selt", "Select top 100", "SELECT TOP 100 *\r\nFROM #\r\nWHERE 1=1"),
                new SnippetItem("upd", "Update template", "UPDATE #\r\nSET \r\nWHERE 1=1"),
                new SnippetItem("del", "Delete template", "DELETE FROM #\r\nWHERE 1=1"),
                new SnippetItem("ins", "Insert template", "INSERT INTO # ()\r\nVALUES ()"),
                new SnippetItem("beg", "Begin/End block", "BEGIN\r\n    #\r\nEND"),
                new SnippetItem("iff", "If block", "IF #\r\nBEGIN\r\n    \r\nEND"),
                new SnippetItem("trycatch", "Try/Catch block", "BEGIN TRY\r\n    #\r\nEND TRY\r\nBEGIN CATCH\r\n    SELECT ERROR_MESSAGE() AS ErrorMessage\r\nEND CATCH"),
                new SnippetItem("cte", "Common Table Expression", "WITH CTE AS (\r\n    SELECT #\r\n)\r\nSELECT * FROM CTE"),
                new SnippetItem("tempt", "Temp table", "CREATE TABLE #TempTable (\r\n    Id INT,\r\n    #\r\n)"),
                new SnippetItem("header", "Script header", "-- =============================================\r\n-- Author:      $USER$\r\n-- Create date: $DATE$\r\n-- Description: #\r\n-- ============================================="),
            };
        }
    }
}
