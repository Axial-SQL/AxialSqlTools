// DatabaseScripterToolWindowControl.xaml.cs
using DocumentFormat.OpenXml.Math;
using Microsoft.Data.SqlClient;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Sdk.Sfc;
using Microsoft.SqlServer.Management.Smo;
using Newtonsoft.Json;
using Octokit;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using static AxialSqlTools.ScriptFactoryAccess;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView;

namespace AxialSqlTools
{
    public partial class DatabaseScripterToolWindowControl : UserControl
    {

        // In-memory list of DB connections
        private readonly ObservableCollection<ConnectionInfo> _connections;
        private ConnectionInfo SelectedConnection => (ConnectionInfo)ConnectionsListBox.SelectedItem;


        private readonly ObservableCollection<GitRepo> _repos;
        private GitRepo SelectedRepo => (GitRepo)ReposListBox.SelectedItem;
        private readonly ObservableCollection<string> _progressMessages = new ObservableCollection<string>();

        public DatabaseScripterToolWindowControl()
        {
            InitializeComponent();
            ProgressListBox.ItemsSource = _progressMessages;

            _connections = new ObservableCollection<ConnectionInfo>();
            ConnectionsListBox.ItemsSource = _connections;

            _repos = new ObservableCollection<GitRepo>(RepoStore.Load());
            ReposListBox.ItemsSource = _repos;
            if (_repos.Any()) ReposListBox.SelectedIndex = 0;
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

            try
            {
                // 1) Script every database into one combined dictionary
                var allScripts = new Dictionary<string, string>();
                int total = _connections.Count;
                for (int i = 0; i < total; i++)
                {
                    var ci = _connections[i];
                    textProgress.Report($"▶ [{i + 1}/{total}] Scripting {ci.ServerName}\\{ci.Database}…");

                    var scripts = await Task.Run(() =>
                        ScriptAllObjectsInMemory(
                            ci.FullConnectionString,
                            ci.Database,
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
                textProgress.Report($"▶ Committing all {allScripts.Count} scripts to {SelectedRepo.DisplayName}…");

                await CommitToGitHubAsync(
                    SelectedRepo,
                    allScripts,
                    MessageTextBox.Text.Trim(),
                    textProgress);

                _progressMessages.Add("🎉 All done!");
            }
            catch (Exception ex)
            {
                _progressMessages.Add("❌ Error: " + ex.Message);
            }
            finally
            {
                RunButton.IsEnabled = true;
            }
        }

        private Dictionary<string, string> ScriptAllObjectsInMemory(
            string connectionString,
            string databaseName,
            IProgress<string> msgProgress,
            IProgress<double> pctProgress)
        {
            // Prepare SMO objects
            var sqlConn = new SqlConnection(connectionString);
            var serverConn = new ServerConnection(sqlConn);
            // server already passed in but reaffirm it's the same:
            Server server = new Server(serverConn);

            // grab the server instance name
            string serverName = server.ConnectionContext.ServerInstance;

            var options = new ScriptingOptions
            {
                IncludeHeaders = false,
                IncludeIfNotExists = false,
                ScriptSchema = true,
                DriAll = true,
                Indexes = true,
                SchemaQualify = true,
                ScriptBatchTerminator = true,
                IncludeDatabaseContext = false
            };
            var scripter = new Scripter(server) { Options = options };

            // map sys.objects types → SMO Urn type + folder + whether to pretty‐format
            var objectMap = new[]
            {
                new { SysType="U",   UrnType="Table",               Folder="Tables",           Format=true  },
                new { SysType="V",   UrnType="View",                Folder="Views",            Format=false },
                new { SysType="P",   UrnType="StoredProcedure",     Folder="StoredProcedures", Format=false },
                new { SysType="FN",  UrnType="UserDefinedFunction", Folder="Functions",        Format=false },
                new { SysType="TF",  UrnType="UserDefinedFunction", Folder="Functions",        Format=false },
                new { SysType="IF",  UrnType="UserDefinedFunction", Folder="Functions",        Format=false },
                new { SysType="TR",  UrnType="DatabaseDdlTrigger",  Folder="Triggers",         Format=false },
                new { SysType="SN",  UrnType="Synonym",             Folder="Synonyms",         Format=false },
                new { SysType="TT",  UrnType="UserDefinedType",     Folder="Types",            Format=false }
            };

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
                    using (var cmd = new SqlCommand(
                        @"SELECT SCHEMA_NAME(schema_id) AS SchemaName, name
                          FROM sys.objects
                          WHERE type=@t AND is_ms_shipped=0
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

                                // ---- HERE: include the server in the URN ----
                                var urnText =
                                    $"Server[@Name='{serverName}']/" +
                                    $"Database[@Name='{databaseName}']/" +
                                    $"{m.UrnType}[@Schema='{schema}' and @Name='{name}']";

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
            string commitMessage,
            IProgress<string> progress)
        {
            var client = new GitHubClient(new ProductHeaderValue("AxialSqlTools"))
            {
                Credentials = new Credentials(repo.Token)
            };

            progress.Report("Fetching branch info...");
            var reference = await client.Git.Reference.Get(repo.Owner, repo.Name, $"heads/{repo.Branch}");
            var latestCommit = await client.Git.Commit.Get(repo.Owner, repo.Name, reference.Object.Sha);

            //------test-
            var remotePaths = new List<string>();
            foreach (var folder in new[] { "Tables", "Views", "StoredProcedures", /* etc */ })
            {
                // this returns only the files & sub-dirs under `<dbName>/<folder>`
                var items = await client.Repository.Content
                                     .GetAllContentsByRef(repo.Owner, repo.Name, $"{"AWSInventory"}/{folder}", repo.Branch);

                // files get `Type == "file"` and you get their .Path and .Sha directly
                remotePaths.AddRange(items
                    .Where(i => i.Type == ContentType.File)
                    .Select(i => i.Path));
            }
            var remoteSet = new HashSet<string>(remotePaths);
            //\\ ---test-

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
            var msg = $"About to commit changes to GitHub:\n\n" +
                      $"Files to add:    {toAdd.Count}\n" +
                      //$"Files to modify: {toModify.Count}\n" +
                      $"Files to delete (doesn't work yet): {toDelete.Count}\n\n" +
                      "Continue?";
            if (MessageBox.Show(msg, "Confirm Commit", MessageBoxButton.YesNo, MessageBoxImage.Question)
                != MessageBoxResult.Yes)
            {
                progress.Report("Commit cancelled by user.");
                return;
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

            //// deletions
            //foreach (var path in toDelete)
            //{
            //    treeItems.Add(new NewTreeItem { Path = path, Mode = "100644", Type = TreeType.Blob, Sha = null });
            //    progress.Report($"[{++count}/{total}] marked delete for {path}");
            //}

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
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
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
            var dir = Path.GetDirectoryName(_path);
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            File.WriteAllText(_path, JsonConvert.SerializeObject(repos, Formatting.Indented));
        }
    }
}
