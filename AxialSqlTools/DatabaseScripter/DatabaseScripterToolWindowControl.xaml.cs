using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Sdk.Sfc;
using Microsoft.SqlServer.Management.Smo;
using Octokit;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace AxialSqlTools
{
    public partial class DatabaseScripterToolWindowControl : UserControl
    {
        private readonly ObservableCollection<string> _progressMessages = new ObservableCollection<string>();

        public DatabaseScripterToolWindowControl()
        {

            InitializeComponent();

            ProgressListBox.ItemsSource = _progressMessages;



        }

        private async void RunButton_Click(object sender, RoutedEventArgs e)
        {
            RunButton.IsEnabled = false;
            _progressMessages.Clear();

            var progress = new Progress<string>(msg =>
            {
                _progressMessages.Add(msg);
                ProgressListBox.ScrollIntoView(msg);
            });

            // 1) Retrieve connection info and SMO objects on the UI thread
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

            try
            {
                // 1. Script everything into memory
                var scripts = await Task.Run(() => ScriptAllObjectsInMemory(server, db, messageProgress, percentProgress));

                // 2. Commit to GitHub
                await CommitToGitHubAsync(
                    owner: OwnerTextBox.Text.Trim(),
                    repo: RepoTextBox.Text.Trim(),
                    branch: BranchTextBox.Text.Trim(),
                    message: MessageTextBox.Text.Trim(),
                    files: scripts,
                    progress: progress);

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
                // only Tables get formatted
                (db.Tables.Cast<Table>()
                            .Where(t => !t.IsSystemObject)
                            .Select(t => t.Urn),
                    "Tables", true),

                (db.Views.Cast<View>()
                            .Where(v => !v.IsSystemObject)
                            .Select(v => v.Urn),
                    "Views", false),

                (db.StoredProcedures.Cast<StoredProcedure>()
                                        .Where(sp => !sp.IsSystemObject)
                                        .Select(sp => sp.Urn),
                    "StoredProcedures", false),

                (db.UserDefinedFunctions.Cast<UserDefinedFunction>()
                                            .Where(fn => !fn.IsSystemObject)
                                            .Select(fn => fn.Urn),
                    "TableValuedFunctions", false),

                (db.Synonyms.Cast<Synonym>()
                            .Select(s => s.Urn),
                    "Synonyms", false),

                (db.UserDefinedDataTypes.Cast<UserDefinedDataType>()
                                            .Select(udt => udt.Urn),
                    "Types", false),

                (db.Triggers.Cast<DatabaseDdlTrigger>()
                            .Where(tr => !tr.IsSystemObject)
                            .Select(tr => tr.Urn),
                    "Triggers", false)
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

                    if (format)
                        sql = TSqlFormatter.FormatCode(sql);

                    var path = $"{folder}/{urn.GetAttribute("Schema")}.{urn.GetAttribute("Name")}.sql";

                    result[path] = sql;

                    done++;

                    msgProgress.Report($"[{done}/{total}] scripted {path}");
                    pctProgress.Report(done * 100.0 / total);

                }
            }

            return result;
        }

        private async Task CommitToGitHubAsync(
            string owner,
            string repo,
            string branch,
            string message,
            Dictionary<string, string> files,
            IProgress<string> progress)
        {
            // load token however you like
            var token = WindowsCredentialHelper.LoadToken("AxialSqlTools_GitHubToken");
            if (string.IsNullOrEmpty(token))
                throw new InvalidOperationException("GitHub token not found");

            var client = new GitHubClient(new ProductHeaderValue("AxialSqlTools"))
            {
                Credentials = new Credentials(token)
            };

            // 1. get the branch ref and latest commit
            progress.Report("Fetching branch info...");
            var reference = await client.Git.Reference.Get(owner, repo, $"heads/{branch}");
            var latestCommit = await client.Git.Commit.Get(owner, repo, reference.Object.Sha);

            // 2. create blobs
            var treeItems = new List<NewTreeItem>();
            int count = 0, total = files.Count;
            foreach (var kvp in files)
            {
                var blob = await client.Git.Blob.Create(owner, repo,
                    new NewBlob { Content = kvp.Value, Encoding = EncodingType.Utf8 });

                treeItems.Add(new NewTreeItem
                {
                    Path = kvp.Key,
                    Mode = "100644",
                    Type = TreeType.Blob,
                    Sha = blob.Sha
                });

                count++;
                progress.Report($"[{count}/{total}] created blob for {kvp.Key}");
            }

            // 3. create new tree
            progress.Report("Creating tree...");
            var newTree = new NewTree { BaseTree = latestCommit.Tree.Sha };
            foreach (var ti in treeItems) newTree.Tree.Add(ti);
            var treeResponse = await client.Git.Tree.Create(owner, repo, newTree);

            // 4. commit
            progress.Report("Creating commit...");
            var newCommit = await client.Git.Commit.Create(owner, repo,
                new NewCommit(message, treeResponse.Sha, new[] { latestCommit.Sha }));

            // 5. update branch
            await client.Git.Reference.Update(owner, repo, reference.Ref,
                new ReferenceUpdate(newCommit.Sha));

            progress.Report($"Commit complete: {newCommit.Sha}");
        }
    }
}
