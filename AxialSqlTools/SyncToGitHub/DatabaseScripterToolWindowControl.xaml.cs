// DatabaseScripterToolWindowControl.xaml.cs
using Microsoft.Data.SqlClient;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Sdk.Sfc;
using Microsoft.SqlServer.Management.Smo;
using Newtonsoft.Json;
using Octokit;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
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


        private readonly ObservableCollection<GitRepo> _repos;
        private GitRepo SelectedRepo => (GitRepo)ReposComboBox.SelectedItem;
        private readonly ObservableCollection<string> _progressMessages = new ObservableCollection<string>();

        public DatabaseScripterToolWindowControl()
        {
            InitializeComponent();
            ProgressListBox.ItemsSource = _progressMessages;

            _connections = new ObservableCollection<ScriptFactoryAccess.ConnectionInfo>();
            ConnectionsListBox.ItemsSource = _connections;

            _repos = new ObservableCollection<GitRepo>(RepoStore.Load());
            ReposComboBox.ItemsSource = _repos;
            // if (_repos.Any()) ReposComboBox.SelectedIndex = 0;
        }

        private void ReposComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            // Only if you haven’t already selected one
            if (ReposComboBox.Items.Count > 0 && ReposComboBox.SelectedIndex < 0)
            {
                // Option A: by index
                ReposComboBox.SelectedIndex = 0;

                // Option B: by item
                // ReposComboBox.SelectedItem = _repos[0];
            }
        }

        private void OpenRepoCommitsButton_Click(object sender, RoutedEventArgs e)
        {
            var repo = SelectedRepo;
            if (repo == null)
            {
                MessageBox.Show("Please select a repo first.", "No Repo Selected",
                                MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // URL to the commits for the current branch
            var url = $"https://github.com/{repo.Owner}/{repo.Name}/commits/{repo.Branch}";

            // Open in default browser
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
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

        private void AddRepoButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new AddRepoDialog
            {
                Owner = Window.GetWindow(this),
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };
            if (dlg.ShowDialog() == true)
            {
                _repos.Add(dlg.Result);
                RepoStore.Save(_repos);
            }
        }

        private void RemoveRepoButton_Click(object sender, RoutedEventArgs e)
        {
            var repo = SelectedRepo;
            if (repo == null)
                return;

            var answer = MessageBox.Show(
                $"Remove {repo.DisplayName}?",
                "Confirm",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (answer != MessageBoxResult.Yes)
                return;

            _repos.Remove(repo);
            RepoStore.Save(_repos);
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

            if (SelectedRepo == null)
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
                textProgress.Report($"▶ Committing all {allScripts.Count} scripts to {SelectedRepo.DisplayName}...");

                await CommitToGitHubAsync(
                    SelectedRepo,
                    allScripts,
                    ObjectMap,
                    MessageTextBox.Text.Trim(),
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

        public static async Task<List<string>> GetAllRemoteFilePathsAsync(
            GitHubClient client,
            GitRepo repo,
            string basePath)
        {
            var result = new List<string>();
            await Recurse(basePath).ConfigureAwait(false);
            return result;

            async Task Recurse(string path)
            {
                System.Collections.Generic.IReadOnlyList<RepositoryContent> items;
                try
                {
                    items = await client.Repository.Content
                        .GetAllContentsByRef(repo.Owner, repo.Name, path, repo.Branch)
                        .ConfigureAwait(false);
                }
                catch (NotFoundException)
                {
                    // directory doesn't exist → nothing under here
                    return;
                }

                foreach (var item in items)
                {
                    if (item.Type == ContentType.Dir)
                    {
                        // dive into subfolder
                        await Recurse(item.Path).ConfigureAwait(false);
                    }
                    else if (item.Type == ContentType.File)
                    {
                        result.Add(item.Path);
                    }
                }
            }
        }

        private Dictionary<string, string> ScriptAllObjectsInMemory(
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

            return result;
        }


        private async Task CommitToGitHubAsync(
            GitRepo repo,
            Dictionary<string, string> files,
            SmoObjectMap[] objectMap,
            string commitMessage,
            IProgress<string> progress)
        {
            var client = new GitHubClient(new ProductHeaderValue("AxialSqlTools"))
            {
                Credentials = new Credentials(repo.Token)
            };

            // Prepare a list of files to be deleted
            var remotePaths = new List<string>();
            foreach (ScriptFactoryAccess.ConnectionInfo ci in _connections)
            {
                progress.Report($"▶ Fetching remote file structure under [{ci.Database}]");
                var items = await GetAllRemoteFilePathsAsync(client, SelectedRepo, ci.Database);
                remotePaths.AddRange(items);
            }                
            var remoteSet = new HashSet<string>(remotePaths);
            //------------

            var localPaths = new HashSet<string>(files.Keys);
            //var toAdd = localPaths.Except(remotePaths).ToList();
            var toAdd = localPaths.ToList();
            var toDelete = remotePaths.Except(localPaths).ToList();
            //var toModify = localPaths
            //               .Intersect(remotePaths)
            //               .Where(p => {
            //                   var remoteContent = FetchRemoteContent(client, repo, p).GetAwaiter().GetResult();
            //                   return !string.Equals(remoteContent, files[p], StringComparison.Ordinal);
            //               })
            //               .ToList();

            // confirmation dialog

            if (ConfirmBeforePushing.IsChecked.GetValueOrDefault()) {

                var msg = $"About to commit changes to GitHub:\n\n" +
                            $"Files to add/modify:    {toAdd.Count}\n" +
                            //$"Files to modify: {toModify.Count}\n" +
                            $"Files to delete: {toDelete.Count}\n\n" +
                            "Continue?";
                if (MessageBox.Show(msg, "Confirm Commit", MessageBoxButton.YesNo, MessageBoxImage.Question)
                    != MessageBoxResult.Yes)
                {
                    progress.Report("Commit cancelled by user.");
                    return;
                }
            }

            // deletions
            int delCounter = 0, delTotal = toDelete.Count;
            foreach (var path in toDelete)
            {
                // // there is a bug here - https://github.com/octokit/octokit.net/issues/2836
                // treeItems.Add(new NewTreeItem { Path = path, Mode = "100644", Type = TreeType.Blob, Content = "<object deleted>"});
                // progress.Report($"[{++count}/{total}] emptied file for {path}");

                progress.Report($"[{++delCounter}/{delTotal}] Deleting {path} ... ");

                var existing = await client.Repository.Content
                    .GetAllContentsByRef(repo.Owner, repo.Name, path, repo.Branch)
                    .ConfigureAwait(false);
                var fileSha = existing.First().Sha;

                // 2) Build your DeleteFileRequest
                var deleteRequest = new DeleteFileRequest(
                    message: $"Remove {path}",
                    sha: fileSha,
                    branch: repo.Branch
                // optional: author, committer can be set too
                );

                // 3) Call the DeleteFile method
                await client.Repository.Content.DeleteFile(repo.Owner, repo.Name, path, deleteRequest).ConfigureAwait(false);

            }


            // build tree items
            var treeItems = new List<NewTreeItem>();
            int count = 0, total = toAdd.Count; // + toModify.Count + toDelete.Count;

            // additions & modifications
            foreach (var path in toAdd) // toAdd.Concat(toModify))
            {
                var blob = await client.Git.Blob.Create(repo.Owner, repo.Name,
                    new NewBlob { Content = files[path], Encoding = EncodingType.Utf8 });
                treeItems.Add(new NewTreeItem { Path = path, Mode = "100644", Type = TreeType.Blob, Sha = blob.Sha });
                progress.Report($"[{++count}/{total}] prepared blob for {path}");
            }           

            progress.Report("Fetching branch info...");
            var reference = await client.Git.Reference.Get(repo.Owner, repo.Name, $"heads/{repo.Branch}");
            var latestCommit = await client.Git.Commit.Get(repo.Owner, repo.Name, reference.Object.Sha);

            // create new tree & commit
            progress.Report("Creating tree...");
            var newTree = new NewTree { BaseTree = latestCommit.Tree.Sha };
            foreach (var ti in treeItems) newTree.Tree.Add(ti);
            var treeResponse = await client.Git.Tree.Create(repo.Owner, repo.Name, newTree);

            progress.Report("Creating commit...");
            var newCommit = await client.Git.Commit.Create(repo.Owner, repo.Name,
                new NewCommit(commitMessage, treeResponse.Sha, new[] { latestCommit.Sha }));

            progress.Report("Updating branch...");
            await client.Git.Reference.Update(repo.Owner, repo.Name, reference.Ref,
                new ReferenceUpdate(newCommit.Sha));

            progress.Report($"Commit complete: {newCommit.Sha}");
        }

        private async Task<string> FetchRemoteContent(GitHubClient client, GitRepo repo, string path)
        {
            var contents = await client.Repository.Content.GetAllContentsByRef(repo.Owner, repo.Name, path, repo.Branch);
            return contents.First().Content;
        }
    }

    // --- helper classes below ---

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
    public static class RepoStore
    {
        private static readonly string _path =
            System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "AxialSQL", "repos.json");

        public static List<GitRepo> Load()
        {
            try
            {
                if (!File.Exists(_path)) return new List<GitRepo>();
                var json = File.ReadAllText(_path);
                return JsonConvert.DeserializeObject<List<GitRepo>>(json);
            }
            catch { return new List<GitRepo>(); }
        }

        public static void Save(IEnumerable<GitRepo> repos)
        {
            var dir = System.IO.Path.GetDirectoryName(_path);
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            File.WriteAllText(_path, JsonConvert.SerializeObject(repos, Formatting.Indented));
        }
    }
}
