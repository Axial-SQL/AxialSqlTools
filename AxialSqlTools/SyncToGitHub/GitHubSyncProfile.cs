using AxialSqlTools;
using Microsoft.SqlServer.Management.Common;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

public class GitHubSyncProfile
{
    public string ProfileName { get; set; }
    public GitRepo Repo { get; set; } = new GitRepo();

    public string DatabaseList { get; set; }

    [JsonProperty("ServerConnection")]
    public string EncryptedServerConnection { get; set; }

    [JsonIgnore]
    private ScriptFactoryAccess.ConnectionInfo _serverConnection;

    [JsonIgnore]
    public ScriptFactoryAccess.ConnectionInfo ServerConnection
    {
        get
        {
            if (_serverConnection == null && !string.IsNullOrEmpty(EncryptedServerConnection))
            {
                try
                {
                    var cipher = Convert.FromBase64String(EncryptedServerConnection);
                    var plainBytes = SettingsManager.Unprotect(cipher);
                    if (plainBytes != null)
                    {
                        var json = Encoding.UTF8.GetString(plainBytes);
                        _serverConnection = JsonConvert.DeserializeObject<ScriptFactoryAccess.ConnectionInfo>(json);
                    }
                    else
                    {
                        _serverConnection = null;
                    }
                }
                catch
                {
                    _serverConnection = null;
                }
            }
            return _serverConnection;
        }
        set
        {
            _serverConnection = value;

            if (_serverConnection != null)
            {
                var json = JsonConvert.SerializeObject(_serverConnection);
                var data = Encoding.UTF8.GetBytes(json);
                var cipher = SettingsManager.Protect(data);
                EncryptedServerConnection = Convert.ToBase64String(cipher);
            }
            else
            {
                EncryptedServerConnection = null;
            }
        }
    }

    public bool ExportServerConfigValues { get; set; }
    public bool ExportServerJobs { get; set; }
    public bool ExportServerLoginsAndPermissions { get; set; }

    public override string ToString() => ProfileName;
}
