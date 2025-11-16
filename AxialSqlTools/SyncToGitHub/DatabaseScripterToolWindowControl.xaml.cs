using Microsoft.Data.SqlClient;
using Microsoft.SqlServer.Dac;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Sdk.Sfc;
using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.Management.Smo.Agent;
using Octokit;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using static AxialSqlTools.DatabaseScripterToolWindowControl;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView;

namespace AxialSqlTools
{
    public partial class DatabaseScripterToolWindowControl : UserControl
    {
        // ripped out of Ola script using AI - https://github.com/olahallengren/sql-server-maintenance-solution/
        private string _dbPullQuery = @"
SET NOCOUNT ON;

DECLARE @norm nvarchar(max) = ISNULL(@Templates, N'');
SET @norm = REPLACE(REPLACE(@norm, CHAR(9), N''), N' ', N'');

-- XML-safe
SET @norm = REPLACE(REPLACE(REPLACE(REPLACE(@norm, N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;'), N'""', N'&quot;');

DECLARE @Tokens TABLE
(
	token nvarchar(4000) NOT NULL,
	is_exclude bit NOT NULL,
	is_wildcard bit NOT NULL
);

DECLARE @xml xml = N'<x><v>' + REPLACE(@norm, N',', N'</v><v>') + N'</v></x>';

INSERT @Tokens(token, is_exclude, is_wildcard)
SELECT
	CASE WHEN LEFT(T.Item, 1) = '-' THEN STUFF(T.Item, 1, 1, N'') ELSE T.Item END,
	CASE WHEN LEFT(T.Item, 1) = '-' THEN 1 ELSE 0 END,
	CASE WHEN CHARINDEX(N'%', T.Item) > 0 OR CHARINDEX(N'_', T.Item) > 0 THEN 1 ELSE 0 END
FROM (
	SELECT v.value('.', 'nvarchar(4000)') AS Item
	FROM @xml.nodes('/x/v') AS N(v)
) AS T
WHERE LEN(T.Item) > 0;

-- Normalize [DbName]
UPDATE @Tokens
  SET token = CASE WHEN LEFT(token,1)='[' AND RIGHT(token,1)=']'
				   THEN SUBSTRING(token, 2, LEN(token)-2) ELSE token END;

-- Reusable materialized sets
DECLARE @sysdb TABLE(name sysname PRIMARY KEY);
INSERT @sysdb(name)
SELECT d.name
FROM sys.databases AS d
WHERE d.name <> N'tempdb'
  AND d.source_database_id IS NULL;  -- exclude snapshots

DECLARE @agdb TABLE(database_name sysname PRIMARY KEY);
INSERT @agdb(database_name)
SELECT DISTINCT adc.database_name
FROM sys.availability_databases_cluster AS adc;

-- Include expansion
DECLARE @Include TABLE (database_name sysname PRIMARY KEY);

INSERT INTO @Include(database_name)
SELECT DISTINCT s.name
FROM @Tokens AS t
CROSS APPLY (
	SELECT s2.name
	FROM @sysdb AS s2
	WHERE
		-- Keywords
		(t.is_exclude = 0 AND t.token IN (N'ALL_DATABASES', N'USER_DATABASES', N'SYSTEM_DATABASES', N'AVAILABILITY_GROUP_DATABASES')
		 AND (
				(t.token = N'ALL_DATABASES')
			 OR (t.token = N'USER_DATABASES' AND s2.name NOT IN (N'master', N'model', N'msdb'))
			 OR (t.token = N'SYSTEM_DATABASES' AND s2.name IN (N'master', N'model', N'msdb'))
			 OR (t.token = N'AVAILABILITY_GROUP_DATABASES' AND EXISTS (SELECT 1 FROM @agdb g WHERE g.database_name = s2.name))
			 )
		)
		OR
		-- Literal include
		(t.is_exclude = 0 AND t.token NOT IN (N'ALL_DATABASES', N'USER_DATABASES', N'SYSTEM_DATABASES', N'AVAILABILITY_GROUP_DATABASES')
		 AND t.is_wildcard = 0 AND s2.name = t.token)
		OR
		-- Wildcard include
		(t.is_exclude = 0 AND t.is_wildcard = 1 AND s2.name LIKE t.token ESCAPE '\')
) AS s(name);

-- Exclude expansion
DECLARE @ExcludeNames TABLE (database_name sysname PRIMARY KEY);
DECLARE @ExcludePatterns TABLE (pattern nvarchar(4000));

-- Exclude keywords
INSERT INTO @ExcludeNames(database_name)
SELECT DISTINCT s2.name
FROM @Tokens AS t
JOIN @sysdb AS s2 ON 1 = 1
WHERE t.is_exclude = 1
  AND t.token IN (N'ALL_DATABASES', N'USER_DATABASES', N'SYSTEM_DATABASES', N'AVAILABILITY_GROUP_DATABASES')
  AND (
		(t.token = N'ALL_DATABASES')
	 OR (t.token = N'USER_DATABASES' AND s2.name NOT IN (N'master', N'model', N'msdb'))
	 OR (t.token = N'SYSTEM_DATABASES' AND s2.name IN (N'master', N'model', N'msdb'))
	 OR (t.token = N'AVAILABILITY_GROUP_DATABASES' AND EXISTS (SELECT 1 FROM @agdb g WHERE g.database_name = s2.name))
  );

-- Exclude literals
INSERT INTO @ExcludeNames(database_name)
SELECT DISTINCT t.token
FROM @Tokens AS t
WHERE t.is_exclude = 1
  AND t.token NOT IN (N'ALL_DATABASES', N'USER_DATABASES', N'SYSTEM_DATABASES', N'AVAILABILITY_GROUP_DATABASES')
  AND t.is_wildcard = 0;

-- Exclude patterns
INSERT INTO @ExcludePatterns(pattern)
SELECT DISTINCT t.token
FROM @Tokens AS t
WHERE t.is_exclude = 1
  AND t.is_wildcard = 1;

-- Finalize: includes minus excludes

SELECT i.database_name
FROM @Include AS i
WHERE NOT EXISTS (SELECT 1 FROM @ExcludeNames AS e WHERE e.database_name = i.database_name)
  AND NOT EXISTS (
		SELECT 1
		FROM @ExcludePatterns AS p
		WHERE i.database_name LIKE p.pattern ESCAPE '\'
  );
";
        public class SmoObjectMap
        {
            public string SysType { get; }
            public string UrnType { get; }
            public string Folder { get; }
            public bool Format { get; }
            public string ParentType { get; }  

            public SmoObjectMap(
                string sysType,
                string urnType,
                string folder,
                bool format,
                string parentType = null)
            {
                SysType = sysType;
                UrnType = urnType;
                Folder = folder;
                Format = format;
                ParentType = parentType;
            }
        }        

        private readonly ObservableCollection<string> _progressMessages = new ObservableCollection<string>();

        private readonly ObservableCollection<GitHubSyncProfile> _profiles;
        private GitHubSyncProfile _currentProfile;
        private bool _isEditingProfile = false;

        private ScriptFactoryAccess.ConnectionInfo SelectedProfileConnection;

        private void buttonWikiPage_Click(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
        }


        #region SyncProfiles UI

        private void UpdateMainFrameState()
        {
            bool hasProfile = ProfilesComboBox.SelectedItem != null;

            // Disable controls below profile selector if no profile selected
            RunButton.IsEnabled = hasProfile;
            MessageTextBox.IsEnabled = hasProfile;
            ConfirmBeforePushing.IsEnabled = hasProfile;
        }

        private void UpdateRepoInfo(GitHubSyncProfile profile)
        {
            if (profile == null || profile.Repo == null)
            {
                RepoUrlHyperlink.Inlines.Clear();
                RepoUrlHyperlink.Inlines.Add("Repo: (none selected)");
                RepoUrlHyperlink.NavigateUri = null;

                RepoSyncOptionsTextBlock.Text = "Sync options will appear here";
                return;
            }

            // Build the GitHub repo URL
            string url = $"https://github.com/{profile.Repo.Owner}/{profile.Repo.Name}/tree/{profile.Repo.Branch}";

            RepoUrlHyperlink.Inlines.Clear();
            RepoUrlHyperlink.Inlines.Add(url);
            RepoUrlHyperlink.NavigateUri = new Uri(url);

            var jobsInfo = profile.ExportServerJobs ? "✔ Export SQL Agent Parameters" : "✖ SQL Agent Parameters Not Exported";
            var loginsInfo = profile.ExportServerLoginsAndPermissions ? "✔ Export Logins" : "✖ Logins Not Exported";
            var serverConfInfo = profile.ExportServerConfigValues ? "✔ Export Server Config Values" : "✖ Server Config Not Exported";
            RepoSyncOptionsTextBlock.Text = $"{serverConfInfo} | {jobsInfo} | {loginsInfo}";

            var targetDbInfo = profile.DatabaseList;
            RepoSyncDbsTextBlock.Text = $"Target Databases: {targetDbInfo}";
        }

        private void RepoUrlHyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            if (e.Uri != null)
            {
                Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            }
            e.Handled = true;
        }



        // Switch to profile editor
        private void AddNewProfileButton_Click(object sender, RoutedEventArgs e)
        {
            _currentProfile = new GitHubSyncProfile();
            _isEditingProfile = false;

            // Clear all fields
            ProfileNameTextBox.Clear();
            RepoOwnerTextBox.Clear();
            RepoNameTextBox.Clear();
            RepoBranchTextBox.Text = "main";
            RepoTokenBox.Password = string.Empty;

            SelectedProfileConnection = null;
            ProfileConnectionInfoTextBox.Text = "No server connection selected";
            ListOfDatabasesTextBox.Text = "";

            ExportServerParamsCheckBox.IsChecked = false;
            ExportJobsCheckBox.IsChecked = false;
            ExportLoginsCheckBox.IsChecked = false;

            ProfileEditorFrame.Visibility = Visibility.Visible;
            MainFrame.Visibility = Visibility.Collapsed;
        }

        private void SetProfileServerConnection_Click(object sender, RoutedEventArgs e)
        {

            SelectedProfileConnection = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer(inMaster : true);
            ProfileConnectionInfoTextBox.Text = SelectedProfileConnection.ToString();
        }

        private void EditProfileButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedProfile = (GitHubSyncProfile)ProfilesComboBox.SelectedItem;
            if (selectedProfile == null)
            {
                MessageBox.Show("Please select a profile to edit.", "No Profile Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _currentProfile = selectedProfile;
            _isEditingProfile = true;

            // Pre-fill form fields
            ProfileNameTextBox.Text = _currentProfile.ProfileName;

            if (_currentProfile.Repo != null)
            {
                RepoOwnerTextBox.Text = _currentProfile.Repo.Owner;
                RepoNameTextBox.Text = _currentProfile.Repo.Name;
                RepoBranchTextBox.Text = _currentProfile.Repo.Branch;
                RepoTokenBox.Password = _currentProfile.Repo.Token ?? string.Empty;
            }

            SelectedProfileConnection = _currentProfile.ServerConnection;
            if (SelectedProfileConnection != null)
            {
                ProfileConnectionInfoTextBox.Text = SelectedProfileConnection.ToString();
            }
            ListOfDatabasesTextBox.Text = _currentProfile.DatabaseList ?? string.Empty;

            ExportServerParamsCheckBox.IsChecked = _currentProfile.ExportServerConfigValues;
            ExportJobsCheckBox.IsChecked = _currentProfile.ExportServerJobs;
            ExportLoginsCheckBox.IsChecked = _currentProfile.ExportServerLoginsAndPermissions;

            ProfileEditorFrame.Visibility = Visibility.Visible;
            MainFrame.Visibility = Visibility.Collapsed;
        }

        // Save profile
        private async void SaveProfile_Click(object sender, RoutedEventArgs e)
        {
            string profileName = ProfileNameTextBox.Text.Trim();
            string owner = RepoOwnerTextBox.Text.Trim();
            string name = RepoNameTextBox.Text.Trim();
            string branch = RepoBranchTextBox.Text.Trim();
            string token = RepoTokenBox.Password.Trim();

            // Basic validation
            if (string.IsNullOrEmpty(profileName) ||
                string.IsNullOrEmpty(owner) ||
                string.IsNullOrEmpty(name) ||
                string.IsNullOrEmpty(branch) ||
                string.IsNullOrEmpty(token))
            {
                MessageBox.Show("Profile Name, Owner, Repo Name, Branch, and Token are required.",
                                "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // GitHub repo & branch validation
            try
            {
                var client = new GitHubClient(new ProductHeaderValue("AxialSqlTools"))
                {
                    Credentials = new Credentials(token)
                };

                // Validate repo existence
                var repoInfo = await client.Repository.Get(owner, name);

                // Validate branch existence
                var branchInfo = await client.Repository.Branch.Get(owner, name, branch);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to validate GitHub details:\n{ex.Message}",
                                "GitHub Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Save profile after successful validation
            _currentProfile.ProfileName = profileName;

            if (_currentProfile.Repo == null)
                _currentProfile.Repo = new GitRepo();
            _currentProfile.Repo.Owner = owner;
            _currentProfile.Repo.Name = name;
            _currentProfile.Repo.Branch = branch;
            _currentProfile.Repo.Token = token;      
            _currentProfile.ServerConnection = SelectedProfileConnection;
            _currentProfile.DatabaseList = ListOfDatabasesTextBox.Text.Trim();
            _currentProfile.ExportServerConfigValues = ExportServerParamsCheckBox.IsChecked == true;
            _currentProfile.ExportServerJobs = ExportJobsCheckBox.IsChecked == true;
            _currentProfile.ExportServerLoginsAndPermissions = ExportLoginsCheckBox.IsChecked == true;

            if (!_isEditingProfile)
                _profiles.Add(_currentProfile);

            ProfileStore.Save(_profiles);

            _isEditingProfile = false;
            SwitchToMainFrame();

        }


        private void CancelProfileEdit_Click(object sender, RoutedEventArgs e)
        {
            _isEditingProfile = false;
            SwitchToMainFrame();
        }

        private void DeleteProfile_Click(object sender, RoutedEventArgs e)
        {
            if (_currentProfile == null)
            {
                MessageBox.Show("No profile selected.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var confirm = MessageBox.Show($"Delete profile '{_currentProfile.ProfileName}'?", "Confirm Delete",
                                          MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (confirm == MessageBoxResult.Yes)
            {
                _profiles.Remove(_currentProfile);
                ProfileStore.Save(_profiles);
                _currentProfile = null;
                _isEditingProfile = false;
                SwitchToMainFrame();
            }
        }


        private void SwitchToMainFrame()
        {
            ProfileEditorFrame.Visibility = Visibility.Collapsed;
            MainFrame.Visibility = Visibility.Visible;
            ProfileNameTextBox.Clear();
        }

        // Load profile into main UI
        private void ProfilesComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _currentProfile = (GitHubSyncProfile)ProfilesComboBox.SelectedItem;

            UpdateRepoInfo(_currentProfile);
            UpdateMainFrameState();
        }
        #endregion

        public DatabaseScripterToolWindowControl()
        {
            InitializeComponent();
            ProgressListBox.ItemsSource = _progressMessages;

            _profiles = new ObservableCollection<GitHubSyncProfile>(ProfileStore.Load());
            ProfilesComboBox.ItemsSource = _profiles;

            UpdateMainFrameState();
        }

        private async void RunButton_Click(object sender, RoutedEventArgs e)
        {

            if (_currentProfile.Repo == null)
            {
                MessageBox.Show("Please select a GitHub repo first.", "No Repo Selected",
                                MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            RunButton.IsEnabled = false;
            _progressMessages.Clear();

            IProgress<string> textProgress = new Progress<string>(msg =>
            {
                _progressMessages.Add(msg);
                ProgressListBox.ScrollIntoView(msg);
            });
            IProgress<double> overallProgress = new Progress<double>(pct =>
            {
                OverallProgressBar.Value = pct;
            });

            SmoObjectMap[] ObjectMap =
            {
                new SmoObjectMap("U",  "Table",               "Tables",           true),
                new SmoObjectMap("V",  "View",                "Views",            false),
                new SmoObjectMap("P",  "StoredProcedure",     "StoredProcedures", false),
                new SmoObjectMap("FN", "UserDefinedFunction", "Functions",        false),
                new SmoObjectMap("TF", "UserDefinedFunction", "Functions",        false),
                new SmoObjectMap("IF", "UserDefinedFunction", "Functions",        false),
                //new SmoObjectMap("TR", "DatabaseDdlTrigger",  "Triggers",       false), -- weird that Database Triggers are not in sys.objects
                new SmoObjectMap("TR", "Trigger",             "Triggers",         false, parentType: "Table"), 
                new SmoObjectMap("SN", "Synonym",             "Synonyms",         false),
                new SmoObjectMap("TT", "UserDefinedType",     "Types",            false),
            };

            try
            {
                // 1) Script every database into one combined dictionary
                var allScripts = new Dictionary<string, string>();

                var targetServerConnection = _currentProfile.ServerConnection;

                string serverName = "";
                List<String> allDatabases = new List<string>();

                // - compile list of databases to script using Ola concept
                using (var dbConn = new SqlConnection(targetServerConnection.FullConnectionString))
                {
                    await dbConn.OpenAsync();
                    using (var cmd = new SqlCommand(_dbPullQuery, dbConn))
                    {
                        cmd.Parameters.AddWithValue("@Templates", _currentProfile.DatabaseList);

                        var reader = cmd.ExecuteReader();
                        while (reader.Read())
                        {
                            allDatabases.Add(reader.GetString(0));
                        }
                        reader.Close();
                    }
                    using (var cmd = new SqlCommand("SELECT CONVERT(sysname, SERVERPROPERTY('ServerName'))", dbConn))
                    {
                        serverName = (string)(cmd.ExecuteScalar() ?? "");
                        // named instance
                        serverName = serverName.Replace(@"\", "_");
                    }
                    if (string.IsNullOrWhiteSpace(serverName))
                        serverName = "UnknownServer";
                }

                foreach (string databaseName in allDatabases)
                {
                    textProgress.Report($"▶ Scripting [{targetServerConnection.ServerName}] \\ {databaseName}...");

                    var scripts = await Task.Run(() =>
                        ScriptAllObjectsInMemory(
                            _currentProfile,
                            targetServerConnection.FullConnectionString,
                            serverName,
                            databaseName,
                            ObjectMap,
                            textProgress,
                            // no-op percent reporter
                            new Progress<double>(_ => { })
                        ));

                    // merge all together
                    foreach (var kv in scripts)
                    {
                        allScripts[kv.Key] = kv.Value;
                    }

                }
                

                var additionalScripts = await Task.Run(() =>
                    ScriptAdditionalSettingsInMemory(
                        _currentProfile,
                        targetServerConnection.FullConnectionString,
                        serverName,
                        "master",
                        ObjectMap,
                        textProgress,
                        // no-op percent reporter
                        new Progress<double>(_ => { })
                    ));

                // merge all together
                foreach (var kv in additionalScripts)
                {
                    allScripts[kv.Key] = kv.Value;
                }

                // 2) Commit everything at once to the selected repo
                //textProgress.Report($"▶ Committing all {allScripts.Count} scripts to {SelectedRepo.DisplayName}...");

                bool shouldConfirm = ConfirmBeforePushing.IsChecked.GetValueOrDefault();

                await CommitToGitHubAsync(
                    _currentProfile.Repo,
                    allDatabases,
                    allScripts,
                    ObjectMap,
                    MessageTextBox.Text.Trim(),
                    shouldConfirm,
                    textProgress);

                textProgress.Report("🎉 All done!");
            }
            catch (Exception ex)
            {
                textProgress.Report("❌ Error: " + ex.Message);
            }
            finally
            {
                RunButton.IsEnabled = true;
            }
        }

        private static string SanitizeFileName(string name)
        {
            // Replace invalid filename chars with underscore
            var invalidChars = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(name.Length);
            foreach (var c in name)
            {
                sb.Append(invalidChars.Contains(c) ? '_' : c);
            }

            // Trim trailing dots/spaces
            var result = sb.ToString().TrimEnd(' ');
          
            return result;
        }

        public static async Task<Dictionary<string, string>> GetAllRemoteFileShasAsync(
            GitHubClient client,
            GitRepo repo,
            string basePath,
            IProgress<string> progress)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            async Task Recurse(string path)
            {
                System.Collections.Generic.IReadOnlyList<RepositoryContent> items;
                try
                {

                    progress.Report($"▶ Fetching folder structure under [{path}]");

                    items = await client.Repository.Content
                        .GetAllContentsByRef(repo.Owner, repo.Name, path, repo.Branch)
                        .ConfigureAwait(false);
                }
                catch (NotFoundException)
                {
                    return; // folder doesn’t exist
                }

                foreach (var item in items)
                {
                    if (item.Type == ContentType.Dir)
                    {
                        await Recurse(item.Path).ConfigureAwait(false);
                    }
                    else if (item.Type == ContentType.File)
                    {
                        result[item.Path] = item.Sha;
                    }
                }
            }

            await Recurse(basePath).ConfigureAwait(false);
            return result;
        }

        private Dictionary<string, string> ScriptAllObjectsInMemory(
            GitHubSyncProfile currentProfile,
            string connectionString,
            string serverName,
            string databaseName,
            SmoObjectMap[] objectMap,
            IProgress<string> msgProgress,
            IProgress<double> pctProgress)
        {

            var dacResults = new Dictionary<string, string>();

            string tempRoot = Path.Combine(Path.GetTempPath(), "dac_extract_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempRoot);

            try
            {
                var dac = new DacServices(connectionString);

                var extractOptions = new DacExtractOptions
                {
                    // Matches: /p:ExtractTarget=SchemaObjectType
                    ExtractTarget = DacExtractTarget.SchemaObjectType,
                    // Matches: /p:VerifyExtraction=false
                    VerifyExtraction = false,
                    // IgnoreUserLoginMappings = true -- doesn't work for some reason..
                };

                dac.Extract(
                    targetPath: tempRoot,     // folder when ExtractTarget is folder-based
                    databaseName: databaseName,
                    applicationName: "Snapshot-" + databaseName,
                    applicationVersion: new Version(1, 0, 0, 0),
                    applicationDescription: "Automated extract",
                    tables: null,                      // or list tuples to include reference data per table
                    extractOptions: extractOptions
                );
                

                var files = Directory.EnumerateFiles(tempRoot, "*.sql", SearchOption.AllDirectories)
                                     .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                                     .ToArray();

                foreach (var file in files)
                {

                    var sql = File.ReadAllText(file).Trim();

                    var scriptPath = file.Replace(tempRoot, "").Replace("\\", "/");

                    if (scriptPath.StartsWith("/Security/"))
                    {
                        if (sql.StartsWith("CREATE LOGIN ["))
                        {
                            //remove autogenerated password
                            sql = Regex.Replace(
                                sql,
                                @"PASSWORD\s*=\s*N?'[^']+'",
                                "PASSWORD = N'<..password..>'",
                                RegexOptions.IgnoreCase);

                        }
                    }

                    var path = $"{serverName}/Databases/{databaseName}{scriptPath}";
                    dacResults[path] = sql;  

                }

                msgProgress.Report($"DacFx export completed");

            }
            catch (Exception ex)
            {
                msgProgress.Report($"[ERROR]: {ex.Message}");
            }
            finally
            {
                // delete temp folder with all files
                try
                {
                    if (Directory.Exists(tempRoot))
                        Directory.Delete(tempRoot, recursive: true);
                }
                catch
                {
                    // best effort cleanup
                }
            }  

            return dacResults;

            //\\--------------------------------
            /*
            var sqlConn = new SqlConnection(connectionString);
            var serverConn = new ServerConnection(sqlConn);
            Server server = new Server(serverConn);

            // grab the server instance name
            string serverName = server.ConnectionContext.TrueName;

            // Prepare SMO objects
            var options = new ScriptingOptions
            {
                IncludeHeaders = false,
                IncludeIfNotExists = false,
                ScriptSchema = true,
                DriAll = true,
                Indexes = true,
                // Triggers = true, -- triggers are scripted separately because SMO doesn't inject GO between table definition and it's trigger
                SchemaQualify = true,
                ScriptBatchTerminator = true,
                IncludeDatabaseContext = false,

                LoginSid = true,
                ExtendedProperties = true,                

                // <— ensure only your options are honored
                EnforceScriptingOptions = true
            };
            var scripter = new Scripter(server) { Options = options };

            // pre‐count for progress
            int total = 0;
            using (var countConn = new SqlConnection(connectionString))
            {
                countConn.Open();
                foreach (var m in objectMap)
                {
                    using (var cmd = new SqlCommand(
                        "SELECT COUNT(*) FROM sys.objects WHERE type=@t AND is_ms_shipped=0",
                        countConn))
                    {
                        cmd.Parameters.AddWithValue("@t", m.SysType);
                        total += (int)cmd.ExecuteScalar();
                    }
                }
            }

            var result = new Dictionary<string, string>();
            int done = 0;

            //---------------------------------------------------------
            // Script database itself
            Database dbObject = server.Databases[databaseName];            
            var dbLines = dbObject.Script(options);
            var dbSql = TSqlFormatter.FormatCode(string.Join(Environment.NewLine, dbLines.Cast<string>()));
            var dbPath = $"{serverName}/Databases/{databaseName}/{databaseName}.sql";
            result[dbPath] = dbSql;
            //---------------------------------------------------------

            foreach (var m in objectMap)
            {
                using (var qConn = new SqlConnection(connectionString))
                {
                    qConn.Open();
                    //This approach is WAY faster than the default SMO object iterators
                    using (var cmd = new SqlCommand(
                        $@"USE [{databaseName}]; SELECT SCHEMA_NAME(o.schema_id) AS SchemaName,
                                 o.[name],
                                 SCHEMA_NAME(p.schema_id) AS ParentSchemaName,
                                 p.[name] AS ParentName
                          FROM sys.objects  o
                          LEFT OUTER JOIN sys.objects AS p
                              ON o.parent_object_id = p.[object_id]
                          WHERE o.type=@t AND o.is_ms_shipped=0
                          ORDER BY SchemaName, name",
                        qConn))
                    {
                        cmd.Parameters.AddWithValue("@t", m.SysType);
                        using (var rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                var schema = rdr.GetString(0);
                                var name = rdr.GetString(1);

                                // an example [dbo].["Weird Table Name'];
                                name = name.Replace("'", "''");

                                var urnText =
                                        $"Server[@Name='{serverName}']/" +
                                        $"Database[@Name='{databaseName}']/" +
                                        $"{m.UrnType}[@Schema='{schema}' and @Name='{name}']";

                                if (!string.IsNullOrEmpty(m.ParentType))
                                {
                                    var parentSchema = rdr.GetString(2);
                                    var parentName = rdr.GetString(3);

                                    urnText =
                                        $"Server[@Name='{serverName}']/" +
                                        $"Database[@Name='{databaseName}']/" +
                                        $"{m.ParentType}[@Schema='{parentSchema}' and @Name='{parentName}']/" +
                                        $"{m.UrnType}[@Name='{name}']";

                                }

                                var urn = new Urn(urnText);

                                // script via SMO
                                var lines = scripter.Script(new UrnCollection { urn });
                                var sql = string.Join(Environment.NewLine, lines.Cast<string>());
                                if (m.Format)
                                    sql = TSqlFormatter.FormatCode(sql);

                                var safeKey = SanitizeFileName($"{schema}.{name}.sql");

                                var path = $"{serverName}/Databases/{databaseName}/{m.Folder}/{safeKey}";
                                result[path] = sql;

                                done++;
                                msgProgress.Report($"[{done}/{total}] scripted {path}");
                                pctProgress.Report(done * 100.0 / total);
                            }
                        }
                    }
                }
            }
            
            return result;
            */
        }

        private Dictionary<string, string> ScriptAdditionalSettingsInMemory(
            GitHubSyncProfile currentProfile,
            string connectionString,
            string serverName,
            string databaseName,
            SmoObjectMap[] objectMap,
            IProgress<string> msgProgress,
            IProgress<double> pctProgress)
        {
            // Prepare SMO objects
            var sqlConn = new SqlConnection(connectionString);
            var serverConn = new ServerConnection(sqlConn);
            Server server = new Server(serverConn);

            var options = new ScriptingOptions
            {
                IncludeHeaders = false,
                IncludeIfNotExists = false,
                ScriptSchema = true,
                DriAll = true,
                Indexes = true,
                // Triggers = true, -- triggers are scripted separately because SMO doesn't inject GO between table definition and it's trigger
                SchemaQualify = true,
                ScriptBatchTerminator = true,
                IncludeDatabaseContext = false,

                LoginSid = true,
                ExtendedProperties = true,

                // <— ensure only your options are honored
                EnforceScriptingOptions = true
            };
            var scripter = new Scripter(server) { Options = options };

            var result = new Dictionary<string, string>();            

            if (currentProfile.ExportServerConfigValues)
            {

                var md = new StringBuilder();
                md.AppendLine($"# SQL Server Configuration for `{serverName}`");
                md.AppendLine();
                md.AppendLine("| Name | Configured | Minimum | Maximum | Dynamic | Advanced | Description |");
                md.AppendLine("|------|-----------:|--------:|--------:|:-------:|:--------:|-------------|");

                string Esc(string s) =>
                    string.IsNullOrEmpty(s) ? "" : s.Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");

                foreach (ConfigProperty prop in server.Configuration.Properties)
                {
                    md.AppendLine($"| `{Esc(prop.DisplayName)}` | {prop.ConfigValue} | " +
                                  $"{prop.Minimum} | {prop.Maximum} | {(prop.IsDynamic ? "Yes" : "No")} | " +
                                  $"{(prop.IsAdvanced ? "Yes" : "No")} | {Esc(prop.Description)} |");
                }

                var mdPath = $"{serverName}/ServerConfiguration.md";
                result[mdPath] = md.ToString();

                foreach (LinkedServer ls in server.LinkedServers)
                {
                    var lsScript = ls.Script(options);

                    var safeKey = SanitizeFileName($"{ls.Name}.sql");

                    var path = $"{serverName}/ServerObjects/LinkedServers/{safeKey}";

                    result[path] = string.Join(Environment.NewLine, lsScript.Cast<string>());

                    msgProgress.Report($"[...] scripted Linked Server {ls.Name}");

                }

            }

            if (currentProfile.ExportServerJobs)
            {
                foreach (Job job in server.JobServer.Jobs)
                {
                    var jobScript = job.Script(options);

                    var safeKey = SanitizeFileName($"{job.Name}.sql");

                    var path = $"{serverName}/SQLServerAgent/Jobs/{safeKey}";

                    result[path] = string.Join(Environment.NewLine, jobScript.Cast<string>());

                    msgProgress.Report($"[...] scripted Job [{job.Name}]");
                }

                foreach (Operator agentOperator in server.JobServer.Operators)
                {
                    var safeKey = SanitizeFileName($"{agentOperator.Name}.sql");
                    var opScript = agentOperator.Script(options);
                    var path = $"{serverName}/SQLServerAgent/Operators/{safeKey}";
                    result[path] = string.Join(Environment.NewLine, opScript.Cast<string>());
                    msgProgress.Report($"[...] scripted Operator [{agentOperator.Name}]");
                }

                foreach (Alert alert in server.JobServer.Alerts)
                {
                    var safeKey = SanitizeFileName($"{alert.Name}.sql");
                    var opScript = alert.Script(options);
                    var path = $"{serverName}/SQLServerAgent/Alerts/{safeKey}";
                    result[path] = string.Join(Environment.NewLine, opScript.Cast<string>());
                    msgProgress.Report($"[...] scripted Alert [{alert.Name}]");
                }

                foreach (JobCategory jobCat in server.JobServer.JobCategories)
                {
                    var safeKey = SanitizeFileName($"{jobCat.Name}.sql");
                    var opScript = jobCat.Script(options);
                    var path = $"{serverName}/SQLServerAgent/JobCategories/{safeKey}";
                    result[path] = string.Join(Environment.NewLine, opScript.Cast<string>());
                    msgProgress.Report($"[...] scripted Job Category [{jobCat.Name}]");
                }


            }

            if (currentProfile.ExportServerLoginsAndPermissions)
            {
                //TODO - ideally need to export password hashes here...

                foreach (Login login in server.Logins)
                {
                    var lines = login.Script(options).Cast<string>().ToList();

                    // Replace any "PASSWORD = ..." clause with a static password
                    for (int i = 0; i < lines.Count; i++)
                    {
                        lines[i] = System.Text.RegularExpressions.Regex.Replace(
                            lines[i],
                            @"PASSWORD\s*=\s*[^,]+", // match PASSWORD = <something>
                            "PASSWORD = '<..password..>'" // static replacement
                        );
                    }

                    var loginScript = string.Join(Environment.NewLine, lines);

                    var safeKey = SanitizeFileName($"{login.Name}.sql");

                    var path = $"{serverName}/Security/Logins/{safeKey}";
                    result[path] = loginScript;

                    msgProgress.Report($"[...] scripted Login [{login.Name}]");

                    //// script permissions
                    //foreach (ObjectPermissionInfo perm in login.EnumObjectPermissions())
                    //{
                    //    var permScript = perm.Script();
                    //    var permPath = $"{databaseName}/Logins/{login.Name}_Permissions.sql";
                    //    result[permPath] = string.Join(Environment.NewLine, permScript);
                    //    msgProgress.Report($"[{++done}/{total}] scripted Permissions for {login.Name}");
                    //    pctProgress.Report(done * 100.0 / total);
                    //}
                }

                /* - TODO - No scripting option, do manually
                foreach (Credential cred in server.Credentials)
                {
                    var credScript = cred.Script(options);

                    var path = $"{serverName}/Security/Credentials/{cred.Name}.sql";

                    result[path] = string.Join(Environment.NewLine, credScript.Cast<string>());

                    msgProgress.Report($"[{++done}/{total}] scripted Credential {cred.Name}");
                }
                */
            }

            return result;
        }


        private async Task CommitToGitHubAsync(
            GitRepo repo,
            List<String> allDatabases,
            Dictionary<string, string> files,
            SmoObjectMap[] objectMap,
            string commitMessage,
            bool shouldConfirm,
            IProgress<string> progress)
        {
            var client = new GitHubClient(new ProductHeaderValue("AxialSqlTools"))
            {
                Credentials = new Credentials(repo.Token)
            };


            // 1) Gather all remote path→SHA
            var remoteFileShas = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var topServerFolders = files.Keys
                .Select(p => {
                    var slash = p.IndexOf('/');
                    return slash >= 0 ? p.Substring(0, slash) : p;          // first segment
                })
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var dbSet = new HashSet<string>(allDatabases, StringComparer.OrdinalIgnoreCase);

            foreach (var serverFolder in topServerFolders)
            {
                var shas = await GetAllRemoteFileShasAsync(client, repo, serverFolder, progress)
                                .ConfigureAwait(false);
                foreach (var kv in shas)
                {
                    var parts = kv.Key.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length >= 3)
                    {                       
                        // do not remove database scripts when DB was selected before but not anymore
                        if (parts[1] == "Databases" && !dbSet.Contains(parts[2])) continue;
                        // need to do the same for other export options - agent, server config, etc
                        // TODO

                    }

                    remoteFileShas[kv.Key] = kv.Value;
                }
                    
            }

            var localPaths = new HashSet<string>(files.Keys, StringComparer.OrdinalIgnoreCase);
            var remotePaths = new HashSet<string>(remoteFileShas.Keys, StringComparer.OrdinalIgnoreCase);

            // 2) Diff lists
            var toAdd = localPaths.Except(remotePaths).ToList();
            var toDelete = remotePaths.Except(localPaths).ToList();

            // 3) Check content‐diff for existing files
            var toModify = new List<string>();
            foreach (var path in localPaths.Intersect(remotePaths))
            {
                var localSha = ComputeBlobSha(files[path]);
                var remoteSha = remoteFileShas[path];

                if (!string.Equals(localSha, remoteSha, StringComparison.OrdinalIgnoreCase))
                    toModify.Add(path);
            }

            // 4) Bail out if nothing changed
            var changeCount = toAdd.Count + toModify.Count + toDelete.Count;
            if (changeCount == 0)
            {
                progress.Report("🟢 No changes detected. Skipping GitHub commit.");
                return;
            }

            // 5) Optionally confirm
            if (shouldConfirm)
            {
                var msg = $"About to commit changes to GitHub:\n\n" +
                          $"New files:      {toAdd.Count}\n" +
                          $"Modified files: {toModify.Count}\n" +
                          $"Deleted files:  {toDelete.Count}\n\n" +
                          "Continue?";
                if (MessageBox.Show(msg, "Confirm Commit", MessageBoxButton.YesNo, MessageBoxImage.Question)
                    != MessageBoxResult.Yes)
                {
                    progress.Report("Commit cancelled by user.");
                    return;
                }
            }

            // 6) Delete removed files
            int delCounter = 0;
            foreach (var path in toDelete)
            {
                progress.Report($"[{++delCounter}/{toDelete.Count}] Deleting {path}...");

                // Normalize separators and URL encode
                var apiPath = Uri.EscapeDataString(path);

                // GitHub API expects unescaped forward slashes, so decode them back after encoding
                apiPath = apiPath.Replace("%2F", "/");

                var existing = await client.Repository.Content
                    .GetAllContentsByRef(repo.Owner, repo.Name, apiPath, repo.Branch)
                    .ConfigureAwait(false);

                var fileSha = existing.First().Sha;

                var deleteRequest = new DeleteFileRequest(
                    message: $"Remove {path}",
                    sha: fileSha,
                    branch: repo.Branch
                );

                await client.Repository.Content
                    .DeleteFile(repo.Owner, repo.Name, apiPath, deleteRequest)
                    .ConfigureAwait(false);
            }

            // 7) Prepare blobs only for adds + modifies
            var treeItems = new List<NewTreeItem>();
            int total = toAdd.Count + toModify.Count;

            if (total > 0)
            {
                foreach (var path in toAdd.Concat(toModify))
                {
                    //var blob = await client.Git.Blob.Create(repo.Owner, repo.Name,
                    //    new NewBlob { Content = files[path], Encoding = EncodingType.Utf8 });
                    treeItems.Add(new NewTreeItem
                    {
                        Path = path,
                        Mode = "100644",
                        Type = TreeType.Blob,
                        //Sha = blob.Sha,
                        Content = files[path]
                    });
                    // progress.Report($"[{++count}/{total}] Prepared blob for {path}");
                }

                // 8) Create tree + commit
                progress.Report("Fetching branch info...");
                var reference = await client.Git.Reference
                    .Get(repo.Owner, repo.Name, $"heads/{repo.Branch}")
                    .ConfigureAwait(false);
                var latestCommit = await client.Git.Commit
                    .Get(repo.Owner, repo.Name, reference.Object.Sha)
                    .ConfigureAwait(false);

                progress.Report("Creating tree...");
                var newTree = new NewTree { BaseTree = latestCommit.Tree.Sha };
                foreach (var ti in treeItems)
                    newTree.Tree.Add(ti);
                var treeResponse = await client.Git.Tree
                    .Create(repo.Owner, repo.Name, newTree)
                    .ConfigureAwait(false);

                progress.Report("Creating commit...");
                var newCommit = await client.Git.Commit
                    .Create(repo.Owner, repo.Name,
                            new NewCommit(commitMessage, treeResponse.Sha, new[] { latestCommit.Sha }))
                    .ConfigureAwait(false);

                progress.Report("Updating branch...");
                await client.Git.Reference
                    .Update(repo.Owner, repo.Name, reference.Ref, new ReferenceUpdate(newCommit.Sha))
                    .ConfigureAwait(false);

                progress.Report($"🎉 Commit complete: {newCommit.Sha}");

            }
            else 
            { 
                progress.Report($"🎉 Commit complete");
            }
                
        }

        // helper to compute the Git blob SHA1
        private static string ComputeBlobSha(string content)
        {
            var contentBytes = Encoding.UTF8.GetBytes(content);
            var header = $"blob {contentBytes.Length}\0";
            var headerBytes = Encoding.UTF8.GetBytes(header);

            using (var sha1 = SHA1.Create())
            {
                // feed header + content
                sha1.TransformBlock(headerBytes, 0, headerBytes.Length, null, 0);
                sha1.TransformFinalBlock(contentBytes, 0, contentBytes.Length);
                return BitConverter
                    .ToString(sha1.Hash)
                    .Replace("-", string.Empty)
                    .ToLowerInvariant();
            }
        }

       
    }
    
}
