using Aurora;
using AxialSqlTools;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

internal static partial class CommandRegistrationTests
{
    private static void RunNativePerformanceTests()
    {
        Run("512 sibling folders use one parent index probe each on cold and warm renders", LargeSiblingMenusStayLinear);
        Run("Native popup attachment accepts either verified adjacent insertion slot", AdjacentPopupInsertionIsVerified);
        Run("A popup bound to an unexpected child is rejected without relabeling another control", IncorrectPopupAttachmentIsRejected);
        Run("An unexpected returned native bar is never deleted without verified ownership", UnexpectedReturnedBarIsPreserved);
        Run("Resolved GUI registration preserves overrides and scopes its cached command", ResolvedCommandPreservesVirtualRegistration);
        Run("A native identity change during publication is still rejected", ResolvedCommandStillValidatesNativeBinding);
        Run("512-script production rendering uses 15 UI yields with linear cold and warm lookups", RendererUsesBatchedUiAccess);
    }

    private static QueryTemplateFolder SiblingTree(int count)
    {
        string root = Path.Combine(Path.GetTempPath(), "AxialSqlTools-sibling-performance");
        var folders = new List<QueryTemplateFolder>();
        for (int i = 0; i < count; i++)
        {
            string path = Path.Combine(root, "Folder" + i.ToString("D4"));
            folders.Add(new QueryTemplateFolder(path, new List<QueryTemplateFolder>(),
                new List<QueryTemplateFile> { new QueryTemplateFile(Path.Combine(path, "script.sql")) }));
        }
        return new QueryTemplateFolder(root, folders, new List<QueryTemplateFile>());
    }

    private static void LargeSiblingMenusStayLinear()
    {
        var fixture = new Fixture();
        AddBuiltIns(fixture);
        const int count = 512;
        var tree = SiblingTree(count);
        var active = new HashSet<string>(tree.Folders.Select(folder => folder.Files[0].CommandName));
        for (int pass = 0; pass < 2; pass++)
        {
            if (pass > 0)
            {
                ClearMenu(fixture);
                RemoveUnused(fixture, active);
            }
            fixture.Bar.Controls.IndexReads = 0;
            fixture.App.Commands.ItemCalls = 0;
            fixture.App.Commands.CommandInfoCallCount = 0;
            int previousEnumerations = fixture.Bar.Controls.EnumerationCalls;
            fixture.Bar.Controls.ThrowOnEnumeration = true;
            int checkpoints = 0;
            var warnings = new List<string>();
            var result = QueryTemplateMenuControls.PopulateAsync(tree, fixture.Bar, fixture.Plugin, fixture.Registry,
                file => fixture.Handler(file.CommandName), warnings,
                () => { checkpoints++; return Task.CompletedTask; }, CancellationToken.None).GetAwaiter().GetResult();
            Equal(count, result.Loaded);
            Equal(0, result.Failures);
            Equal(0, warnings.Count);
            Equal(count, fixture.Bar.Controls.IndexReads);
            Equal(count * 2, checkpoints); // One per folder and one per script.
            Equal(count, fixture.App.Commands.CommandInfoCallCount);
            Equal(count * (pass == 0 ? 2 : 1), fixture.App.Commands.ItemCalls);
            Equal(previousEnumerations, fixture.Bar.Controls.EnumerationCalls);
            Equal(count + 2, fixture.Bar.Controls.Count);
            // Directly assigned folder captions still correspond to their native child bars.
            for (int i = 0; i < count; i++)
            {
                var popup = (CommandBarPopup)fixture.Bar.Controls[i + 3];
                Equal(tree.Folders[i].Name, popup.Caption);
                Equal(tree.Folders[i].CommandName, popup.CommandBar.Name);
            }
        }
        Equal(0, fixture.App.Commands.EnumerationCalls);
    }

    private static void AdjacentPopupInsertionIsVerified()
    {
        foreach (bool priorSlot in new[] { false, true })
        {
            var fixture = new Fixture();
            AddBuiltIns(fixture);
            var refresh = fixture.Bar.Controls[1];
            var open = fixture.Bar.Controls[2];
            fixture.Proffer.InsertPopupAtPriorSlot = priorSlot;
            fixture.Bar.Controls.IndexReads = 0;
            var popup = fixture.Plugin.AddCommandBarMenuPopup("AxialSqlTools_Folder_slot", fixture.Bar, appendOnly: true);
            Equal("AxialSqlTools_Folder_slot", popup.CommandBar.Name);
            Equal(priorSlot ? 2 : 1, fixture.Bar.Controls.IndexReads);
            Assert(!refresh.Deleted && !open.Deleted, "The adjacent-slot probe damaged a built-in action.");
        }
    }

    private static void IncorrectPopupAttachmentIsRejected()
    {
        var fixture = new Fixture();
        AddBuiltIns(fixture);
        fixture.Proffer.AttachPopupToWrongBar = true;
        fixture.Bar.Controls.IndexReads = 0;
        Throws<InvalidOperationException>(() => fixture.Plugin.AddCommandBarMenuPopup("AxialSqlTools_Folder_bad", fixture.Bar, appendOnly: true));
        Equal(2, fixture.Bar.Controls.IndexReads);
        Assert(fixture.Proffer.LastCreatedBar.Deleted, "Failed attachment left the newly returned bar registered.");
        Equal("Refresh Templates", fixture.Bar.Controls[1].Caption);
        Equal("Open Templates Folder", fixture.Bar.Controls[2].Caption);
    }

    private static void ResolvedCommandPreservesVirtualRegistration()
    {
        var fixture = new Fixture();
        var handler = new RegistrationOverride(fixture.Plugin, "AxialSqlTools_Template_virtual");
        fixture.App.Commands.ItemCalls = 0;
        fixture.Registry.RegisterCommand(false, handler, true, appendOnly: true);
        Equal(1, handler.Registrations);
        Equal(2, fixture.App.Commands.ItemCalls); // Initial miss plus the newly registered command.
        var command = fixture.Command(handler.CanonicalName);
        fixture.App.Commands.ItemCalls = 0;
        handler.RegisterGUI(command, fixture.Bar, true);
        Equal(2, handler.Registrations);
        Equal(1, fixture.App.Commands.ItemCalls); // Direct public calls must resolve afresh.
        fixture.App.Commands.ItemCalls = 0;
        fixture.Registry.RegisterCommand(false, handler, true);
        Equal(3, handler.Registrations);
        Equal(1, fixture.App.Commands.ItemCalls); // Warm registry call shares the resolved command.
        Equal(1, fixture.Bar.Controls.Count);
    }

    private static void UnexpectedReturnedBarIsPreserved()
    {
        var fixture = new Fixture();
        AddBuiltIns(fixture);
        var unrelated = new CommandBar { Name = "Existing unrelated native bar" };
        fixture.Proffer.ReturnedBarOverride = unrelated;
        Throws<InvalidOperationException>(() => fixture.Plugin.AddCommandBarMenuPopup("AxialSqlTools_Folder_requested", fixture.Bar, appendOnly: true));
        Assert(!unrelated.Deleted, "The host returned an unrelated native bar and cleanup deleted it.");
    }

    private static void ResolvedCommandStillValidatesNativeBinding()
    {
        var fixture = new Fixture();
        const string name = "AxialSqlTools_Template_changed_during_registration";
        fixture.MenuService.WhileAdding = command => fixture.App.Commands.Item(name, -1).ID += 10000;
        Throws<InvalidOperationException>(() => fixture.Registry.RegisterCommand(false, fixture.Handler(name), true, appendOnly: true));
        Equal(0, fixture.Bar.Controls.Count);
        Equal(1, fixture.App.Commands.CommandInfoCallCount);
    }

    private static void RendererUsesBatchedUiAccess()
    {
        var fixture = new Fixture();
        AddBuiltIns(fixture);
        const int count = 512;
        var tree = FlatTree(count);
        var active = new HashSet<string>(tree.Files.Select(file => file.CommandName));
        for (int pass = 0; pass < 2; pass++)
        {
            if (pass > 0)
            {
                ClearMenu(fixture);
                RemoveUnused(fixture, active);
            }
            int gates = 0;
            var batcher = new QueryTemplateUiBatcher(() => true,
                token => { gates++; return Task.CompletedTask; }, () => 0L);
            fixture.App.Commands.ItemCalls = 0;
            fixture.App.Commands.CommandInfoCallCount = 0;
            var warnings = new List<string>();
            var result = QueryTemplateMenuControls.PopulateAsync(tree, fixture.Bar, fixture.Plugin, fixture.Registry,
                file => fixture.Handler(file.CommandName), warnings,
                () => batcher.CheckpointAsync(CancellationToken.None), CancellationToken.None).GetAwaiter().GetResult();
            Equal(count, result.Loaded);
            Equal(0, result.Failures);
            Equal(0, warnings.Count);
            Equal(15, gates);
            Equal(gates, batcher.YieldCount);
            Equal(count, fixture.App.Commands.CommandInfoCallCount);
            Equal(count * (pass == 0 ? 2 : 1), fixture.App.Commands.ItemCalls);
            Equal(count + 2, fixture.Bar.Controls.Count);
        }
    }

    private sealed class RegistrationOverride : CommandBase
    {
        public int Registrations { get; private set; }
        public RegistrationOverride(Plugin plugin, string name) : base(name, name, plugin, name) { }
        public override bool RegisterGUI(OleMenuCommand command, CommandBar bar, bool toolbarOnly, bool appendOnly = false)
        {
            Registrations++;
            return base.RegisterGUI(command, bar, toolbarOnly, appendOnly);
        }
        public override bool IsEnabled() => true;
        public override bool OnCommand() => true;
    }
}
