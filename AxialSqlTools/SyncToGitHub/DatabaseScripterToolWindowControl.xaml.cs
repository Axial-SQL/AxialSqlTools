// DatabaseScripterToolWindowControl.xaml.cs
using Microsoft.Data.SqlClient;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Sdk.Sfc;
using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.Management.Smo.Agent;
using Newtonsoft.Json;
using Octokit;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using static AxialSqlTools.DatabaseScripterToolWindowControl;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView;

namespace AxialSqlTools
{
    public partial class DatabaseScripterToolWindowControl : UserControl
    {
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

        // In-memory list of DB connections
        private readonly ObservableCollection<ScriptFactoryAccess.ConnectionInfo> _connections;
        private ScriptFactoryAccess.ConnectionInfo SelectedConnection => (ScriptFactoryAccess.ConnectionInfo)ConnectionsListBox.SelectedItem;

        private readonly ObservableCollection<string> _progressMessages = new ObservableCollection<string>();

        private readonly ObservableCollection<GitHubSyncProfile> _profiles;
        private GitHubSyncProfile _currentProfile;
        private bool _isEditingProfile = false;


        #region SyncProfiles UI

        private void UpdateMainFrameState()
        {
            bool hasProfile = ProfilesComboBox.SelectedItem != null;

            // Disable controls below profile selector if no profile selected
            ConnectionsListBox.IsEnabled = hasProfile;
            AddConnectionButton.IsEnabled = hasProfile;
            RemoveConnectionButton.IsEnabled = hasProfile;
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

            var jobsInfo = profile.ExportServerJobs ? "✔ Export Jobs" : "✖ Jobs Not Exported";
            var loginsInfo = profile.ExportServerLoginsAndPermissions ? "✔ Export Logins" : "✖ Logins Not Exported";
            RepoSyncOptionsTextBlock.Text = $"{jobsInfo} | {loginsInfo}";
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

            ExportJobsCheckBox.IsChecked = false;
            ExportLoginsCheckBox.IsChecked = false;
            ProfileDatabasesListBox.Items.Clear();

            ProfileEditorFrame.Visibility = Visibility.Visible;
            MainFrame.Visibility = Visibility.Collapsed;
        }

        // Add DB to profile in editor
        private void AddDatabaseToProfile_Click(object sender, RoutedEventArgs e)
        {
            var conn = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer();
            if (conn != null)
            {
                ProfileDatabasesListBox.Items.Add(conn);
            }
        }

        private void RemoveDatabaseFromProfile_Click(object sender, RoutedEventArgs e)
        {
            var selectedDb = (ScriptFactoryAccess.ConnectionInfo)ProfileDatabasesListBox.SelectedItem;
            if (selectedDb != null)
            {
                ProfileDatabasesListBox.Items.Remove(selectedDb);
            }
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
            

            ExportJobsCheckBox.IsChecked = _currentProfile.ExportServerJobs;
            ExportLoginsCheckBox.IsChecked = _currentProfile.ExportServerLoginsAndPermissions;

            ProfileDatabasesListBox.Items.Clear();
            foreach (var db in _currentProfile.Databases)
                ProfileDatabasesListBox.Items.Add(db);

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
            _currentProfile.ExportServerJobs = ExportJobsCheckBox.IsChecked == true;
            _currentProfile.ExportServerLoginsAndPermissions = ExportLoginsCheckBox.IsChecked == true;
            _currentProfile.Databases = ProfileDatabasesListBox.Items.Cast<ScriptFactoryAccess.ConnectionInfo>().ToList();

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
            ProfileDatabasesListBox.Items.Clear();
        }

        // Load profile into main UI
        private void ProfilesComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _currentProfile = (GitHubSyncProfile)ProfilesComboBox.SelectedItem;

            _connections.Clear();
            if (_currentProfile != null)
            {
                foreach (var db in _currentProfile.Databases)
                    _connections.Add(db);
            }

            UpdateRepoInfo(_currentProfile);
            UpdateMainFrameState();
        }
        #endregion

        public DatabaseScripterToolWindowControl()
        {
            InitializeComponent();
            ProgressListBox.ItemsSource = _progressMessages;

            _connections = new ObservableCollection<ScriptFactoryAccess.ConnectionInfo>();
            ConnectionsListBox.ItemsSource = _connections;

            _profiles = new ObservableCollection<GitHubSyncProfile>(ProfileStore.Load());
            ProfilesComboBox.ItemsSource = _profiles;

            UpdateMainFrameState();
        }

        private void AddConnectionButton_Click(object sender, RoutedEventArgs e)
        {
            _connections.Add(ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer());
        }

        private void RemoveConnectionButton_Click(object sender, RoutedEventArgs e)
        {
            var conn = SelectedConnection;
            if (conn == null) return;

            if (MessageBox.Show(
                    $"Remove connection to {conn.ServerName}\\{conn.Database}?",
                    "Confirm",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning)
                == MessageBoxResult.Yes)
            {
                _connections.Remove(conn);
            }
        }

        private async void RunButton_Click(object sender, RoutedEventArgs e)
        {

            if (!_connections.Any())
            {
                MessageBox.Show("Please add at least one database connection first.",
                                "No Connections",
                                MessageBoxButton.OK,
                                MessageBoxImage.Exclamation);
                return;
            }

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
                int total = _connections.Count;
                for (int i = 0; i < total; i++)
                {
                    var ci = _connections[i];
                    textProgress.Report($"▶ [{i + 1}/{total}] Scripting {ci.ServerName}\\{ci.Database}...");

                    var scripts = await Task.Run(() =>
                        ScriptAllObjectsInMemory(
                            _currentProfile,
                            ci.FullConnectionString,
                            ci.Database,
                            ObjectMap,
                            textProgress,
                            // no-op percent reporter
                            new Progress<double>(_ => { })
                        ));

                    // merge, prefixing each script path by the database name
                    foreach (var kv in scripts)
                    { 
                        allScripts[kv.Key] = kv.Value;
                    }
                }

                // 2) Commit everything at once to the selected repo
                //textProgress.Report($"▶ Committing all {allScripts.Count} scripts to {SelectedRepo.DisplayName}...");

                bool shouldConfirm = ConfirmBeforePushing.IsChecked.GetValueOrDefault();

                await CommitToGitHubAsync(
                    _currentProfile.Repo,
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
            string databaseName,
            SmoObjectMap[] objectMap,
            IProgress<string> msgProgress,
            IProgress<double> pctProgress)
        {
            // Prepare SMO objects
            var sqlConn = new SqlConnection(connectionString);
            var serverConn = new ServerConnection(sqlConn);
            // server already passed in but reaffirm it's the same:
            Server server = new Server(serverConn);

            // grab the server instance name
            string serverName = server.ConnectionContext.TrueName;

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

            foreach (var m in objectMap)
            {
                using (var qConn = new SqlConnection(connectionString))
                {
                    qConn.Open();
                    //This approach is WAY faster than the default SMO object iterators
                    using (var cmd = new SqlCommand(
                        @"SELECT SCHEMA_NAME(o.schema_id) AS SchemaName,
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

                                var path = $"{databaseName}/{m.Folder}/{schema}.{name}.sql";
                                result[path] = sql;

                                done++;
                                msgProgress.Report($"[{done}/{total}] scripted {path}");
                                pctProgress.Report(done * 100.0 / total);
                            }
                        }
                    }
                }
            }
            /*
            if (currentProfile.ExportServerJobs)
            {
                foreach (Job job in server.JobServer.Jobs)
                {                    
                    var jobScript = job.Script(options);

                    var path = $"{serverName}/Jobs/{job.Name}.sql";

                    result[path] = string.Join(Environment.NewLine, jobScript.Cast<string>());

                    msgProgress.Report($"[{++done}/{total}] scripted Job {job.Name}");
                }                
            }

            if (currentProfile.ExportServerLoginsAndPermissions)
            {
                foreach (Login login in server.Logins)
                {

                    var loginScript = login.Script(options);
                    var path = $"{serverName}/Logins/{login.Name}.sql";
                    result[path] = string.Join(Environment.NewLine, loginScript.Cast<string>());

                    msgProgress.Report($"[{++done}/{total}] scripted Login {login.Name}");
                    
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
            }
            */
            return result;
        }


        private async Task CommitToGitHubAsync(
            GitRepo repo,
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
            foreach (var ci in _connections)
            {
                // progress.Report($"▶ Fetching remote structure under [{ci.Database}]");
                var shas = await GetAllRemoteFileShasAsync(client, repo, ci.Database, progress)
                                 .ConfigureAwait(false);
                foreach (var kv in shas)
                    remoteFileShas[kv.Key] = kv.Value;
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
                var existing = await client.Repository.Content
                    .GetAllContentsByRef(repo.Owner, repo.Name, path, repo.Branch)
                    .ConfigureAwait(false);
                var fileSha = existing.First().Sha;

                var deleteRequest = new DeleteFileRequest(
                    message: $"Remove {path}",
                    sha: fileSha,
                    branch: repo.Branch
                );
                await client.Repository.Content
                    .DeleteFile(repo.Owner, repo.Name, path, deleteRequest)
                    .ConfigureAwait(false);
            }

            // 7) Prepare blobs only for adds + modifies
            var treeItems = new List<NewTreeItem>();
            int count = 0, total = toAdd.Count + toModify.Count;

            if (total > 0)
            {
                foreach (var path in toAdd.Concat(toModify))
                {
                    var blob = await client.Git.Blob.Create(repo.Owner, repo.Name,
                        new NewBlob { Content = files[path], Encoding = EncodingType.Utf8 });
                    treeItems.Add(new NewTreeItem
                    {
                        Path = path,
                        Mode = "100644",
                        Type = TreeType.Blob,
                        Sha = blob.Sha
                    });
                    progress.Report($"[{++count}/{total}] Prepared blob for {path}");
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
