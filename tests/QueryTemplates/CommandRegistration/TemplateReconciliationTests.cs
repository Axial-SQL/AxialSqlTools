using Aurora;
using AxialSqlTools;
using Microsoft.VisualStudio.CommandBars;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

internal static partial class CommandRegistrationTests
{
    private static void RunReconciliationTests()
    {
        Run("Unchanged refresh preserves root and nested native objects without any native writes", UnchangedReconciliationIsReadOnly);
        Run("Adding one nested script creates only that button and command", ReconciliationAddsOneScript);
        Run("Removing one nested script deletes only that button and command", ReconciliationRemovesOneScript);
        Run("Renaming and moving scripts preserve all unaffected native objects and click paths", ReconciliationRenameAndMove);
        Run("Changing script contents performs no menu mutations and clicks read the new contents", ReconciliationContentEdit);
        Run("Restart reconnects handlers to persisted native controls without rebuilding them", ReconciliationRestart);
        Run("Duplicate scripts and shared-bar popups are removed without damaging retained controls", ReconciliationDuplicateCleanup);
        Run("A failed addition recovers without clearing working controls", ReconciliationAddFailureRecovers);
        Run("A failed obsolete-control deletion recovers without clearing working controls", ReconciliationDeleteFailureRecovers);
        Run("Reconciliation verifies both built-in commands before any mutation", ReconciliationBuiltInPreflight);
        Run("Renaming a folder replaces only its affected branch", ReconciliationFolderRename);
        Run("A case-only rename updates metadata and handler path without replacing the control", ReconciliationCaseOnlyRename);
        Run("An unreadable branch preserves existing controls and handlers", ReconciliationUnreadableBranch);
        Run("A transient native identity failure preserves controls and retries without duplicates", ReconciliationIdentityFailureRecovers);
        Run("An existing empty folder toggles state when its first script is added or removed", ReconciliationEmptyFolderState);
        Run("A native child snapshot failure keeps existing subtree handlers alive", ReconciliationChildSnapshotFailure);
        Run("A successfully identified unbound legacy control does not block a new script", ReconciliationUnboundLegacyControl);
        Run("Large incremental refresh preserves retained controls and inserts a sorted prefix with linear native work", ReconciliationLargeSortedDelta);
    }

    private static void UnchangedReconciliationIsReadOnly()
    {
        using (var test = IncrementalFixture.Create())
        {
            test.RefreshSuccessfully();
            var before = new NativeTreeSnapshot(test.Native);
            foreach (var control in before.Controls.Values) control.ThrowOnTagWrite = true;
            NativeMutationLog.Reset();
            var result = test.RefreshSuccessfully();
            Equal(4, result.Loaded);
            Assert(!result.HasChanges, "An unchanged tree was reported as changed.");
            AssertNoNativeMutations();
            AssertRetained(before, new NativeTreeSnapshot(test.Native));
            test.Click("root.sql", "root original");
            test.Click("Common/a.sql", "a original");
            test.Click("Common/Nested/b.sql", "b original");
            test.Click("Other/c.sql", "c original");
        }
    }

    private static void ReconciliationAddsOneScript()
    {
        using (var test = IncrementalFixture.Create())
        {
            test.RefreshSuccessfully();
            var before = new NativeTreeSnapshot(test.Native);
            test.Write("Common/00-new.sql", "new nested script");
            NativeMutationLog.Reset();
            var result = test.RefreshSuccessfully();
            Equal(5, result.Loaded);
            Equal(1, NativeMutationLog.CommandAdds);
            Equal(1, NativeMutationLog.ControlInsertions);
            Equal(1, NativeMutationLog.PictureWrites);
            Equal(0, NativeMutationLog.ControlDeletes);
            Equal(0, NativeMutationLog.CommandDeletes);
            Equal(0, NativeMutationLog.BarDeletes);
            AssertRetained(before, new NativeTreeSnapshot(test.Native));
            test.Click("Common/00-new.sql", "new nested script");
            test.Click("Common/a.sql", "a original");
        }
    }

    private static void ReconciliationRemovesOneScript()
    {
        using (var test = IncrementalFixture.Create())
        {
            test.RefreshSuccessfully();
            var before = new NativeTreeSnapshot(test.Native);
            string removedName = test.ScriptName("Common/a.sql");
            var removed = before.Controls[removedName];
            var staleHandler = test.Handlers[removedName];
            var staleCommand = test.Native.Command(removedName);
            File.Delete(test.PathFor("Common/a.sql"));
            NativeMutationLog.Reset();
            var result = test.RefreshSuccessfully();
            Equal(3, result.Loaded);
            Equal(1, NativeMutationLog.ControlDeletes);
            Equal(1, NativeMutationLog.CommandDeletes);
            Equal(0, NativeMutationLog.ControlInsertions);
            Equal(0, NativeMutationLog.CommandAdds);
            Equal(0, NativeMutationLog.PictureWrites);
            Equal(0, NativeMutationLog.PropertyWrites);
            Assert(removed.Deleted, "The obsolete control was left attached.");
            AssertRetained(before, new NativeTreeSnapshot(test.Native), removedName);
            staleCommand.Invoke();
            Equal(0, staleHandler.Invocations);
            test.Click("Common/Nested/b.sql", "b original");
        }
    }

    private static void ReconciliationRenameAndMove()
    {
        using (var test = IncrementalFixture.Create())
        {
            test.RefreshSuccessfully();
            foreach (var move in new[]
            {
                new[] { "Common/a.sql", "Common/renamed.sql" },
                new[] { "Common/renamed.sql", "Other/moved.sql" }
            })
            {
                var before = new NativeTreeSnapshot(test.Native);
                string oldName = test.ScriptName(move[0]);
                var oldCommand = test.Native.Command(oldName);
                var oldHandler = test.Handlers[oldName];
                int previousInvocations = oldHandler.Invocations;
                File.Move(test.PathFor(move[0]), test.PathFor(move[1]));
                NativeMutationLog.Reset();
                Equal(4, test.RefreshSuccessfully().Loaded);
                Equal(1, NativeMutationLog.ControlDeletes);
                Equal(1, NativeMutationLog.ControlInsertions);
                Equal(1, NativeMutationLog.CommandDeletes);
                Equal(1, NativeMutationLog.CommandAdds);
                Equal(1, NativeMutationLog.PictureWrites);
                AssertRetained(before, new NativeTreeSnapshot(test.Native), oldName);
                oldCommand.Invoke();
                Equal(previousInvocations, oldHandler.Invocations);
                test.Click(move[1], "a original");
            }
        }
    }

    private static void ReconciliationContentEdit()
    {
        using (var test = IncrementalFixture.Create())
        {
            test.RefreshSuccessfully();
            test.Click("Common/Nested/b.sql", "b original");
            var before = new NativeTreeSnapshot(test.Native);
            test.Write("Common/Nested/b.sql", "select 'new contents';");
            NativeMutationLog.Reset();
            Assert(!test.RefreshSuccessfully().HasChanges, "A content edit was treated as a menu change.");
            AssertNoNativeMutations();
            AssertRetained(before, new NativeTreeSnapshot(test.Native));
            test.Click("Common/Nested/b.sql", "select 'new contents';");
        }
    }

    private static void ReconciliationRestart()
    {
        using (var test = IncrementalFixture.Create())
        {
            test.RefreshSuccessfully();
            var before = new NativeTreeSnapshot(test.Native);
            var previousHandlers = test.Handlers.Values.ToList();
            test.Native = new Fixture(test.Native.App, test.Native.Bar, test.Native.Group);
            test.Handlers.Clear();
            NativeMutationLog.Reset();
            Equal(4, test.RefreshSuccessfully().Loaded);
            AssertNoNativeMutations();
            AssertRetained(before, new NativeTreeSnapshot(test.Native));
            Equal(4, test.Native.MenuService.Count);
            test.Click("root.sql", "root original");
            test.Click("Common/Nested/b.sql", "b original");
            Assert(previousHandlers.All(handler => handler.Invocations == 0), "A persisted control still invoked its previous-session handler.");
        }
    }

    private static void ReconciliationDuplicateCleanup()
    {
        using (var test = IncrementalFixture.Create())
        {
            test.RefreshSuccessfully();
            var before = new NativeTreeSnapshot(test.Native);
            string rootName = test.ScriptName("root.sql");
            var duplicateScript = (CommandBarControl)test.Native.App.Commands.Item(rootName, -1)
                .AddControl(test.Native.Bar, test.Native.Bar.Controls.Count + 1);
            var retainedPopup = (CommandBarPopup)before.Controls[test.FolderName("Common")];
            var duplicatePopup = new CommandBarPopup { Caption = retainedPopup.Caption, CommandBar = retainedPopup.CommandBar };
            test.Native.Bar.Controls.Add(duplicatePopup);
            // A legacy control can share a friendly caption, but not native identity.
            test.Native.App.Commands.Add("Template_99", test.Native.Group, 99999);
            var legacy = (CommandBarControl)test.Native.App.Commands.Item("Template_99", -1)
                .AddControl(test.Native.Bar, test.Native.Bar.Controls.Count + 1);
            legacy.Caption = "root.sql";
            NativeMutationLog.Reset();
            Equal(4, test.RefreshSuccessfully().Loaded);
            Equal(3, NativeMutationLog.ControlDeletes);
            Equal(0, NativeMutationLog.BarDeletes);
            Equal(0, NativeMutationLog.ControlInsertions);
            Equal(0, NativeMutationLog.CommandDeletes);
            Equal(0, NativeMutationLog.CommandAdds);
            Equal(0, NativeMutationLog.PictureWrites);
            Assert(duplicateScript.Deleted && duplicatePopup.Deleted && legacy.Deleted,
                "Duplicate or legacy controls survived reconciliation.");
            Assert(!retainedPopup.CommandBar.Deleted, "Duplicate-popup cleanup deleted the retained shared child bar.");
            AssertRetained(before, new NativeTreeSnapshot(test.Native));
            test.Click("root.sql", "root original");
            test.Click("Common/Nested/b.sql", "b original");
        }
    }

    private static void ReconciliationAddFailureRecovers()
    {
        using (var test = IncrementalFixture.Create())
        {
            test.RefreshSuccessfully();
            var before = new NativeTreeSnapshot(test.Native);
            test.Write("Other/new.sql", "retry succeeded");
            test.Native.App.Commands.ControlResult = FakeProffer.Failure;
            var failed = test.Refresh();
            Assert(failed.Failures > 0, "A native insertion failure was hidden.");
            AssertRetained(before, new NativeTreeSnapshot(test.Native));
            test.Click("Common/a.sql", "a original");
            test.Native.App.Commands.ControlResult = 0;
            NativeMutationLog.Reset();
            Equal(5, test.RefreshSuccessfully().Loaded);
            Equal(1, NativeMutationLog.ControlInsertions);
            Equal(0, NativeMutationLog.ControlDeletes);
            Equal(0, NativeMutationLog.BarDeletes);
            AssertRetained(before, new NativeTreeSnapshot(test.Native));
            test.Click("Other/new.sql", "retry succeeded");
        }
    }

    private static void ReconciliationDeleteFailureRecovers()
    {
        using (var test = IncrementalFixture.Create())
        {
            test.RefreshSuccessfully();
            var before = new NativeTreeSnapshot(test.Native);
            string removedName = test.ScriptName("Common/a.sql");
            var removed = before.Controls[removedName];
            File.Delete(test.PathFor("Common/a.sql"));
            removed.FailDelete = true;
            var failed = test.Refresh();
            Assert(failed.Failures > 0, "A native deletion failure was hidden.");
            AssertRetained(before, new NativeTreeSnapshot(test.Native), removedName);
            test.Click("Common/Nested/b.sql", "b original");
            removed.FailDelete = false;
            NativeMutationLog.Reset();
            Equal(3, test.RefreshSuccessfully().Loaded);
            Equal(1, NativeMutationLog.ControlDeletes);
            Equal(0, NativeMutationLog.ControlInsertions);
            Equal(0, NativeMutationLog.BarDeletes);
            Equal(0, NativeMutationLog.PictureWrites);
            Assert(removed.Deleted, "Retry did not remove the obsolete native control.");
            AssertRetained(before, new NativeTreeSnapshot(test.Native), removedName);
        }
    }

    private static void ReconciliationBuiltInPreflight()
    {
        using (var test = IncrementalFixture.Create())
        {
            test.RefreshSuccessfully();
            var before = new NativeTreeSnapshot(test.Native);
            test.Native.Bar.Controls[2].FailCommandInfo = true;
            test.Write("new.sql", "must not register");
            NativeMutationLog.Reset();
            Throws<InvalidOperationException>(() => test.Refresh());
            AssertNoNativeMutations();
            AssertRetained(before, new NativeTreeSnapshot(test.Native));
        }
    }

    private static void ReconciliationFolderRename()
    {
        using (var test = IncrementalFixture.Create())
        {
            test.RefreshSuccessfully();
            var before = new NativeTreeSnapshot(test.Native);
            var affected = new[] { test.FolderName("Common"), test.FolderName("Common/Nested"),
                test.ScriptName("Common/a.sql"), test.ScriptName("Common/Nested/b.sql") };
            Directory.Move(test.PathFor("Common"), test.PathFor("Renamed"));
            NativeMutationLog.Reset();
            Equal(4, test.RefreshSuccessfully().Loaded);
            Equal(2, NativeMutationLog.CommandAdds);
            Equal(2, NativeMutationLog.CommandDeletes);
            Equal(2, NativeMutationLog.PictureWrites);
            AssertRetained(before, new NativeTreeSnapshot(test.Native), affected);
            test.Click("Renamed/a.sql", "a original");
            test.Click("Renamed/Nested/b.sql", "b original");
            test.Click("Other/c.sql", "c original");
        }
    }

    private static void ReconciliationCaseOnlyRename()
    {
        using (var test = IncrementalFixture.Create())
        {
            test.RefreshSuccessfully();
            var before = new NativeTreeSnapshot(test.Native);
            string name = test.ScriptName("Common/a.sql");
            var retained = (CommandBarButton)before.Controls[name];
            File.Move(test.PathFor("Common/a.sql"), test.PathFor("Common/A.sql"));
            NativeMutationLog.Reset();
            Equal(4, test.RefreshSuccessfully().Loaded);
            Equal(0, NativeMutationLog.ControlDeletes);
            Equal(0, NativeMutationLog.ControlInsertions);
            Equal(0, NativeMutationLog.CommandAdds);
            Equal(0, NativeMutationLog.CommandDeletes);
            Equal(0, NativeMutationLog.PictureWrites);
            var after = new NativeTreeSnapshot(test.Native);
            Assert(ReferenceEquals(retained, after.Controls[name]), "A case-only rename replaced the native button.");
            Equal("A.sql", retained.Caption);
            Equal(test.PathFor("Common/A.sql"), retained.TooltipText);
            AssertRetained(before, after, name);
            test.Click("Common/A.sql", "a original");
        }
    }

    private static void ReconciliationUnreadableBranch()
    {
        using (var test = IncrementalFixture.Create())
        {
            test.RefreshSuccessfully();
            var before = new NativeTreeSnapshot(test.Native);
            var catalog = QueryTemplateCatalog.Load(test.Root, CancellationToken.None,
                new UnreadableFolderFileSystem(test.PathFor("Common")));
            Assert(catalog.Warnings.Count > 0, "The filesystem fixture did not produce an unreadable branch.");
            NativeMutationLog.Reset();
            test.Refresh(catalog);
            AssertNoNativeMutations();
            AssertRetained(before, new NativeTreeSnapshot(test.Native));
            test.Click("Common/a.sql", "a original");
            test.Click("Common/Nested/b.sql", "b original");
            NativeMutationLog.Reset();
            Equal(4, test.RefreshSuccessfully().Loaded);
            AssertNoNativeMutations();
            AssertRetained(before, new NativeTreeSnapshot(test.Native));
        }
    }

    private static void ReconciliationIdentityFailureRecovers()
    {
        using (var test = IncrementalFixture.Create())
        {
            // Both ancestors depend solely on the one script whose identity
            // temporarily cannot be read.
            File.Delete(test.PathFor("Common/a.sql"));
            test.RefreshSuccessfully();
            var before = new NativeTreeSnapshot(test.Native);
            var temporarilyUnknown = before.Controls[test.ScriptName("Common/Nested/b.sql")];
            temporarilyUnknown.FailCommandInfo = true;
            NativeMutationLog.Reset();
            Assert(test.Refresh().Failures > 0, "An unresolved native identity was silently treated as a new script.");
            Assert(before.Controls[test.FolderName("Common")].Enabled
                && before.Controls[test.FolderName("Common/Nested")].Enabled,
                "An incomplete native identity read disabled a retained ancestor popup.");
            AssertNoNativeMutations();
            AssertRetained(before, new NativeTreeSnapshot(test.Native));
            temporarilyUnknown.FailCommandInfo = false;
            NativeMutationLog.Reset();
            Equal(3, test.RefreshSuccessfully().Loaded);
            AssertNoNativeMutations();
            AssertRetained(before, new NativeTreeSnapshot(test.Native));
            test.Click("Common/Nested/b.sql", "b original");
        }
    }

    private static void ReconciliationEmptyFolderState()
    {
        using (var test = IncrementalFixture.Create())
        {
            Directory.CreateDirectory(test.PathFor("Empty"));
            test.RefreshSuccessfully();
            var before = new NativeTreeSnapshot(test.Native);
            string folderName = test.FolderName("Empty");
            var popup = (CommandBarPopup)before.Controls[folderName];
            Assert(popup.Visible && !popup.Enabled, "An empty folder started enabled.");
            test.Write("Empty/new.sql", "first script");
            NativeMutationLog.Reset();
            Equal(5, test.RefreshSuccessfully().Loaded);
            Equal(1, NativeMutationLog.ControlInsertions);
            Equal(0, NativeMutationLog.ControlDeletes);
            var afterAdd = new NativeTreeSnapshot(test.Native);
            Assert(ReferenceEquals(popup, afterAdd.Controls[folderName]) && popup.Enabled,
                "Adding the first script replaced or failed to enable its existing popup.");
            AssertRetained(before, afterAdd, folderName);
            test.Click("Empty/new.sql", "first script");
            File.Delete(test.PathFor("Empty/new.sql"));
            NativeMutationLog.Reset();
            Equal(4, test.RefreshSuccessfully().Loaded);
            Equal(1, NativeMutationLog.ControlDeletes);
            Equal(0, NativeMutationLog.ControlInsertions);
            var afterRemove = new NativeTreeSnapshot(test.Native);
            Assert(ReferenceEquals(popup, afterRemove.Controls[folderName]) && !popup.Enabled,
                "Removing the final script replaced or failed to disable its existing popup.");
            AssertRetained(before, afterRemove, folderName);
        }
    }

    private static void ReconciliationChildSnapshotFailure()
    {
        using (var test = IncrementalFixture.Create())
        {
            File.Delete(test.PathFor("Common/a.sql"));
            test.RefreshSuccessfully();
            var before = new NativeTreeSnapshot(test.Native);
            var branch = ((CommandBarPopup)before.Controls[test.FolderName("Common/Nested")]).CommandBar;
            branch.Controls.FailIndexReads = true;
            NativeMutationLog.Reset();
            Assert(test.Refresh().Failures > 0, "The native child snapshot failure was not reported.");
            Assert(before.Controls[test.FolderName("Common")].Enabled
                && before.Controls[test.FolderName("Common/Nested")].Enabled,
                "An incomplete native child traversal disabled a retained ancestor popup.");
            AssertNoNativeMutations();
            AssertRetained(before, new NativeTreeSnapshot(test.Native));
            // These clicks must retain actual native command dispatch while the
            // branch could not be inspected, not merely preserve inert controls.
            test.Click("Common/Nested/b.sql", "b original");
            test.Click("Other/c.sql", "c original");
            branch.Controls.FailIndexReads = false;
            NativeMutationLog.Reset();
            Equal(3, test.RefreshSuccessfully().Loaded);
            AssertNoNativeMutations();
            AssertRetained(before, new NativeTreeSnapshot(test.Native));
        }
    }

    private static void ReconciliationUnboundLegacyControl()
    {
        using (var test = IncrementalFixture.Create())
        {
            test.RefreshSuccessfully();
            var before = new NativeTreeSnapshot(test.Native);
            var unbound = new CommandBarButton
            {
                Caption = "new.sql",
                NativeGuid = test.Native.Group.ToString("B"),
                NativeId = 0
            };
            test.Native.Bar.Controls.Add(unbound);
            test.Write("new.sql", "new root script");
            NativeMutationLog.Reset();
            Equal(5, test.RefreshSuccessfully().Loaded);
            Assert(unbound.Deleted, "A known-unbound control was retained indefinitely.");
            Equal(1, NativeMutationLog.ControlDeletes);
            Equal(1, NativeMutationLog.ControlInsertions);
            Equal(1, NativeMutationLog.CommandAdds);
            Equal(0, NativeMutationLog.CommandDeletes);
            Equal(0, NativeMutationLog.BarDeletes);
            AssertRetained(before, new NativeTreeSnapshot(test.Native));
            test.Click("new.sql", "new root script");
            test.Click("root.sql", "root original");
        }
    }

    private static void ReconciliationLargeSortedDelta()
    {
        const int retainedCount = 512;
        const int addedCount = 256;
        using (var test = IncrementalFixture.CreateFlat(retainedCount))
        {
            Equal(retainedCount, test.RefreshSuccessfully().Loaded);
            var before = new NativeTreeSnapshot(test.Native);
            test.Native.App.Commands.ItemCalls = 0;
            test.Native.App.Commands.CommandInfoCallCount = 0;
            test.Native.Bar.Controls.IndexReads = 0;
            NativeMutationLog.Reset();
            Equal(retainedCount, test.RefreshSuccessfully().Loaded);
            AssertNoNativeMutations();
            Equal(0, test.Native.App.Commands.ItemCalls);
            Assert(test.Native.App.Commands.CommandInfoCallCount <= 3 * (retainedCount + 2),
                "An unchanged refresh repeatedly queried native command identities.");
            Assert(test.Native.Bar.Controls.IndexReads <= 2 * (retainedCount + 2),
                "An unchanged refresh repeatedly scanned native controls.");
            AssertRetained(before, new NativeTreeSnapshot(test.Native));

            for (int i = 0; i < addedCount; i++)
                test.Write("added-" + i.ToString("D4") + ".sql", "added " + i);
            test.Native.App.Commands.ItemCalls = 0;
            test.Native.App.Commands.CommandInfoCallCount = 0;
            test.Native.Bar.Controls.IndexReads = 0;
            NativeMutationLog.Reset();
            Equal(retainedCount + addedCount, test.RefreshSuccessfully().Loaded);
            Equal(addedCount, NativeMutationLog.ControlInsertions);
            Equal(addedCount, NativeMutationLog.CommandAdds);
            Equal(addedCount, NativeMutationLog.PictureWrites);
            Equal(0, NativeMutationLog.ControlDeletes);
            Equal(0, NativeMutationLog.CommandDeletes);
            Equal(0, NativeMutationLog.BarDeletes);
            Assert(test.Native.App.Commands.ItemCalls <= 2 * addedCount,
                "Adding scripts re-resolved unchanged native commands.");
            Assert(test.Native.App.Commands.CommandInfoCallCount <= 3 * (retainedCount + addedCount + 2),
                "Adding a sorted prefix caused repeated native identity scans.");
            Assert(test.Native.Bar.Controls.IndexReads <= 2 * (retainedCount + addedCount + 2),
                "Adding a sorted prefix caused repeated native index scans.");
            AssertRetained(before, new NativeTreeSnapshot(test.Native));
            for (int i = 0; i < addedCount; i++)
                Equal("added-" + i.ToString("D4") + ".sql", test.Native.Bar.Controls[i + 3].Caption);
            for (int i = 0; i < retainedCount; i++)
                Equal("retained-" + i.ToString("D4") + ".sql", test.Native.Bar.Controls[addedCount + i + 3].Caption);
            test.Click("added-0000.sql", "added 0");
            test.Click("added-0255.sql", "added 255");
            test.Click("retained-0000.sql", "retained 0");
            test.Click("retained-0511.sql", "retained 511");
        }
    }

    private static void AssertNoNativeMutations()
    {
        Equal(0, NativeMutationLog.PropertyWrites);
        Equal(0, NativeMutationLog.PictureWrites);
        Equal(0, NativeMutationLog.StyleWrites);
        Equal(0, NativeMutationLog.ControlInsertions);
        Equal(0, NativeMutationLog.ControlDeletes);
        Equal(0, NativeMutationLog.BarDeletes);
        Equal(0, NativeMutationLog.CommandAdds);
        Equal(0, NativeMutationLog.CommandDeletes);
    }

    private static void AssertRetained(NativeTreeSnapshot before, NativeTreeSnapshot after, params string[] removedNames)
    {
        var removed = new HashSet<string>(removedNames);
        foreach (var entry in before.Controls)
        {
            if (removed.Contains(entry.Key)) continue;
            Assert(after.Controls.TryGetValue(entry.Key, out var current) && ReferenceEquals(entry.Value, current),
                "An unaffected native control was replaced: " + entry.Key);
            Assert(!entry.Value.Deleted, "An unaffected native control was deleted: " + entry.Key);
            Equal(before.PropertyWrites[entry.Key], entry.Value.PropertyWrites);
            if (entry.Value is CommandBarButton button)
                Equal(before.PictureWrites[entry.Key], button.PictureWrites);
        }
        foreach (var entry in before.Bars)
        {
            if (removed.Contains(entry.Key)) continue;
            Assert(after.Bars.TryGetValue(entry.Key, out var current) && ReferenceEquals(entry.Value, current),
                "An unaffected native child bar was replaced: " + entry.Key);
            Assert(!entry.Value.Deleted, "An unaffected native child bar was deleted: " + entry.Key);
        }
    }

    private sealed class NativeTreeSnapshot
    {
        internal readonly Dictionary<string, CommandBarControl> Controls = new Dictionary<string, CommandBarControl>();
        internal readonly Dictionary<string, CommandBar> Bars = new Dictionary<string, CommandBar>();
        internal readonly Dictionary<string, int> PropertyWrites = new Dictionary<string, int>();
        internal readonly Dictionary<string, int> PictureWrites = new Dictionary<string, int>();

        internal NativeTreeSnapshot(Fixture fixture) { Read(fixture, fixture.Bar); }

        private void Read(Fixture fixture, CommandBar bar)
        {
            foreach (CommandBarControl control in bar.Controls)
            {
                string key;
                if (control is CommandBarPopup popup)
                {
                    key = popup.CommandBar.Name;
                    Bars.Add(key, popup.CommandBar);
                    Read(fixture, popup.CommandBar);
                }
                else
                {
                    // Test inspection does not query command status or open a menu.
                    try { key = fixture.App.Commands.GetByIdentity(control.NativeGuid, control.NativeId).Name; }
                    catch (InvalidOperationException) { key = "Unregistered:" + control.NativeGuid + ":" + control.NativeId; }
                    if (control is CommandBarButton button) PictureWrites.Add(key, button.PictureWrites);
                }
                Controls.Add(key, control);
                PropertyWrites.Add(key, control.PropertyWrites);
            }
        }
    }

    private sealed class IncrementalFixture : IDisposable
    {
        internal Fixture Native = new Fixture();
        internal readonly Dictionary<string, FileReadingHandler> Handlers = new Dictionary<string, FileReadingHandler>();
        internal readonly string Root = Path.Combine(Path.GetTempPath(), "AxialSqlTools-reconcile-" + Guid.NewGuid().ToString("N"));
        internal List<string> Warnings;

        private IncrementalFixture() { Directory.CreateDirectory(Root); AddBuiltIns(Native); }

        internal static IncrementalFixture Create()
        {
            var result = new IncrementalFixture();
            result.Write("root.sql", "root original");
            result.Write("Common/a.sql", "a original");
            result.Write("Common/Nested/b.sql", "b original");
            result.Write("Other/c.sql", "c original");
            return result;
        }

        internal static IncrementalFixture CreateFlat(int count)
        {
            var result = new IncrementalFixture();
            for (int i = 0; i < count; i++)
                result.Write("retained-" + i.ToString("D4") + ".sql", "retained " + i);
            return result;
        }

        internal string PathFor(string relative) => Path.Combine(Root, relative.Replace('/', Path.DirectorySeparatorChar));
        internal string ScriptName(string relative) => QueryTemplateCatalog.GetCommandName(PathFor(relative), false);
        internal string FolderName(string relative) => QueryTemplateCatalog.GetCommandName(PathFor(relative), true);

        internal void Write(string relative, string contents)
        {
            string path = PathFor(relative);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, contents);
        }

        internal QueryTemplateReconcileResult Refresh(QueryTemplateCatalog catalog = null)
        {
            catalog = catalog ?? QueryTemplateCatalog.Load(Root, CancellationToken.None);
            Warnings = new List<string>(catalog.Warnings);
            var result = QueryTemplateMenuControls.ReconcileAsync(catalog.Root, Native.Bar, Native.Plugin, Native.Registry,
                file =>
                {
                    var handler = new FileReadingHandler(Native.Plugin, file);
                    Handlers[file.CommandName] = handler;
                    return handler;
                }, RefreshId, OpenFolderId, Warnings, () => Task.CompletedTask, CancellationToken.None).GetAwaiter().GetResult();
            // Use the production reconciliation result for the package's second
            // phase. Never independently derive which handlers should survive.
            try
            {
                Native.Registry.RemoveUnusedCommandsAsync(result.ActiveCommandNames,
                    () => Task.CompletedTask, CancellationToken.None).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                result.Failures++;
                Warnings.Add(ex.Message);
            }
            return result;
        }

        internal QueryTemplateReconcileResult RefreshSuccessfully()
        {
            var result = Refresh();
            Assert(result.Failures == 0, "Reconciliation failed: " + string.Join(" | ", Warnings));
            Equal(0, Warnings.Count);
            Equal(0, Native.App.Commands.EnumerationCalls);
            return result;
        }

        internal void Click(string relative, string expectedContents)
        {
            string name = ScriptName(relative);
            var button = (CommandBarButton)new NativeTreeSnapshot(Native).Controls[name];
            var handler = Handlers[name];
            int previous = handler.Invocations;
            button.Click();
            Equal(previous + 1, handler.Invocations);
            Equal(PathFor(relative), handler.FilePath);
            Equal(expectedContents, handler.LastContents);
        }

        public void Dispose() { Directory.Delete(Root, true); }
    }

    private sealed class FileReadingHandler : CommandBase
    {
        internal readonly string FilePath;
        internal int Invocations;
        internal string LastContents;

        internal FileReadingHandler(Plugin plugin, QueryTemplateFile file)
            : base(file.Name, file.CommandName, plugin, file.FullPath) { FilePath = file.FullPath; }

        public override bool IsEnabled() => true;
        public override bool OnCommand()
        {
            LastContents = File.ReadAllText(FilePath);
            Invocations++;
            return true;
        }
    }

    private sealed class UnreadableFolderFileSystem : IQueryTemplateFileSystem
    {
        private readonly string _unreadable;
        internal UnreadableFolderFileSystem(string unreadable) { _unreadable = unreadable; }
        public string[] GetDirectories(string path)
        {
            if (path == _unreadable) throw new UnauthorizedAccessException("Temporary templates share access failure.");
            return Directory.GetDirectories(path);
        }
        public string[] GetFiles(string path) => Directory.GetFiles(path);
        public FileAttributes GetAttributes(string path) => File.GetAttributes(path);
    }
}
