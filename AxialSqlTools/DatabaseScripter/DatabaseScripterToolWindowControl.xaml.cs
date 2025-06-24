// DatabaseScripterToolWindowControl.xaml.cs
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView;

namespace AxialSqlTools
{
    public partial class DatabaseScripterToolWindowControl : UserControl
    {
        private readonly ObservableCollection<GitRepo> _repos;
        private GitRepo SelectedRepo => (GitRepo)ReposListBox.SelectedItem;
        private readonly ObservableCollection<string> _progressMessages = new ObservableCollection<string>();

        public DatabaseScripterToolWindowControl()
        {
            InitializeComponent();
            ProgressListBox.ItemsSource = _progressMessages;

            _repos = new ObservableCollection<GitRepo>(RepoStore.Load());
            ReposListBox.ItemsSource = _repos;
            if (_repos.Any()) ReposListBox.SelectedIndex = 0;
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
            if (SelectedRepo == null)
            {
                MessageBox.Show("Please select a GitHub repo first.", "No Repo Selected",
                                MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            RunButton.IsEnabled = false;
            _progressMessages.Clear();

            var progress = new Progress<string>(msg =>
            {
                _progressMessages.Add(msg);
                ProgressListBox.ScrollIntoView(msg);
            });

            try
            {
                // 1) Script everything into memory
                var ci = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer();
                var server = new Server(new ServerConnection(ci.ServerName));
                var db = server.Databases[ci.Database];

                var messageProgress = new Progress<string>(msg =>
                {
                    _progressMessages.Add(msg);
                    ProgressListBox.ScrollIntoView(msg);
                });
                var percentProgress = new Progress<double>(pct =>
                {
                    OverallProgressBar.Value = pct;
                });

                var scripts = await Task.Run(() =>
                    ScriptAllObjectsInMemory(server, db, messageProgress, percentProgress));

                // 2) Commit to GitHub with diff & confirm
                await CommitToGitHubAsync(
                    SelectedRepo,
                    scripts,
                    MessageTextBox.Text.Trim(),
                    progress);

                _progressMessages.Add("✅ All done!");
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
            Server server,
            Database db,
            IProgress<string> msgProgress,
            IProgress<double> pctProgress)
        {
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

            var result = new Dictionary<string, string>();
            var categories = new (IEnumerable<Urn> items, string folder, bool format)[]
            {
                (db.Tables.Cast<Table>().Where(t=>!t.IsSystemObject).Select(t=>t.Urn), "Tables", true),
                (db.Views.Cast<View>().Where(v=>!v.IsSystemObject).Select(v=>v.Urn), "Views", false),
                (db.StoredProcedures.Cast<StoredProcedure>().Where(sp=>!sp.IsSystemObject).Select(sp=>sp.Urn), "StoredProcedures", false),
                (db.UserDefinedFunctions.Cast<UserDefinedFunction>().Where(fn=>!fn.IsSystemObject).Select(fn=>fn.Urn), "TableValuedFunctions", false),
                (db.Synonyms.Cast<Synonym>().Select(s=>s.Urn), "Synonyms", false),
                (db.UserDefinedDataTypes.Cast<UserDefinedDataType>().Select(udt=>udt.Urn), "Types", false),
                (db.Triggers.Cast<DatabaseDdlTrigger>().Where(tr=>!tr.IsSystemObject).Select(tr=>tr.Urn), "Triggers", false)
            };

            var scripter = new Scripter(server) { Options = options };
            int total = categories.Sum(c => c.items.Count());
            int done = 0;

            foreach (var (items, folder, format) in categories)
            {
                foreach (var urn in items)
                {
                    var lines = scripter.Script(new UrnCollection { urn });
                    var sql = string.Join(Environment.NewLine, lines.Cast<string>());
                    if (format) sql = TSqlFormatter.FormatCode(sql);

                    var path = $"{db.Name}/{folder}/{urn.GetAttribute("Schema")}.{urn.GetAttribute("Name")}.sql";
                    result[path] = sql;

                    done++;
                    msgProgress.Report($"[{done}/{total}] scripted {path}");
                    pctProgress.Report(done * 100.0 / total);
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
