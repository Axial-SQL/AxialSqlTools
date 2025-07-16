using AxialSqlTools;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

public class GitHubSyncProfile
{
    public string ProfileName { get; set; }
    public GitRepo Repo { get; set; } = new GitRepo();

    // Encrypted JSON representation of databases (stored in file)
    [JsonProperty("Databases")]
    public string EncryptedDatabases { get; set; }

    // Backing field for decrypted databases
    [JsonIgnore]
    private List<ScriptFactoryAccess.ConnectionInfo> _databases;

    [JsonIgnore]
    public List<ScriptFactoryAccess.ConnectionInfo> Databases
    {
        get
        {
            if (_databases == null && !string.IsNullOrEmpty(EncryptedDatabases))
            {
                try
                {
                    var cipher = Convert.FromBase64String(EncryptedDatabases);
                    var plainBytes = SettingsManager.Unprotect(cipher);
                    if (plainBytes != null)
                    {
                        var json = Encoding.UTF8.GetString(plainBytes);
                        _databases = JsonConvert.DeserializeObject<List<ScriptFactoryAccess.ConnectionInfo>>(json);
                    }
                    else
                    {
                        _databases = new List<ScriptFactoryAccess.ConnectionInfo>();
                    }
                }
                catch
                {
                    _databases = new List<ScriptFactoryAccess.ConnectionInfo>();
                }
            }
            return _databases ?? new List<ScriptFactoryAccess.ConnectionInfo>();
        }
        set
        {
            _databases = value ?? new List<ScriptFactoryAccess.ConnectionInfo>();

            if (_databases.Count > 0)
            {
                var json = JsonConvert.SerializeObject(_databases);
                var data = Encoding.UTF8.GetBytes(json);
                var cipher = SettingsManager.Protect(data);
                EncryptedDatabases = Convert.ToBase64String(cipher);
            }
            else
            {
                EncryptedDatabases = null;
            }
        }
    }

    public bool ExportServerJobs { get; set; }
    public bool ExportServerLoginsAndPermissions { get; set; }

    public override string ToString() => ProfileName;
}
