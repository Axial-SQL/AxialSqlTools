using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

public static class ProfileStore
{
    private static readonly string _path =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                     "AxialSQL", "github-sync-profiles.json");

    public static List<GitHubSyncProfile> Load()
    {
        try
        {
            if (!File.Exists(_path)) return new List<GitHubSyncProfile>();
            var json = File.ReadAllText(_path);
            return JsonConvert.DeserializeObject<List<GitHubSyncProfile>>(json);
        }
        catch { return new List<GitHubSyncProfile>(); }
    }

    public static void Save(IEnumerable<GitHubSyncProfile> profiles)
    {
        var dir = Path.GetDirectoryName(_path);
        if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
        File.WriteAllText(_path, JsonConvert.SerializeObject(profiles, Formatting.Indented));
    }
}
