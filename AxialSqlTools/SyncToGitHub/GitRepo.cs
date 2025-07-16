using AxialSqlTools;
using Newtonsoft.Json;
using System;
using System.Text;

public class GitRepo
{
    // This is what we persist to disk
    [JsonProperty("Token")]
    public string EncryptedToken { get; set; }

    // Backing field for the decrypted token
    [JsonIgnore]
    private string _token;

    // This is what the rest of your code uses
    [JsonIgnore]
    public string Token
    {
        get
        {
            if (_token == null && !string.IsNullOrEmpty(EncryptedToken))
            {
                // decrypt on first access
                var cipher = Convert.FromBase64String(EncryptedToken);
                var plain = SettingsManager.Unprotect(cipher);
                _token = plain != null
                    ? Encoding.UTF8.GetString(plain)
                    : throw new InvalidOperationException("Failed to decrypt GitRepo token.");
            }
            return _token;
        }
        set
        {
            _token = value;
            if (!string.IsNullOrEmpty(value))
            {
                // encrypt each time it’s set
                var data = Encoding.UTF8.GetBytes(value);
                var cipher = SettingsManager.Protect(data);
                EncryptedToken = Convert.ToBase64String(cipher);
            }
            else
            {
                EncryptedToken = null;
            }
        }
    }

    // other props...
    public string Owner { get; set; }
    public string Name { get; set; }
    public string Branch { get; set; }

    [JsonIgnore]
    public string DisplayName => $"{Owner}/{Name}@{Branch}";
}