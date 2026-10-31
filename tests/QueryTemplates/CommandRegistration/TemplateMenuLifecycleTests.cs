using AxialSqlTools;
using Microsoft.VisualStudio.CommandBars;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

internal static partial class CommandRegistrationTests
{
    private static readonly Guid BuiltInGroup = new Guid("45457e02-6dec-4a4d-ab22-c9ee126d23c5");
    private static readonly CommandID RefreshId = new CommandID(BuiltInGroup, 4132);
    private static readonly CommandID OpenFolderId = new CommandID(BuiltInGroup, 4145);

    private static void RunMenuLifecycleTests()
    {
        Run("Unsupported Tag reproduces the old cleanup miss; native cleanup removes stale controls", UnsupportedTagRegression);
        Run("Cleanup verifies both native built-in actions before any mutation", BuiltInPreflight);
        Run("Transient command lookup after preflight cannot delete either built-in action", BuiltInIdentityTransientFailure);
        Run("Cleanup preserves built-in identities despite order and caption changes", BuiltInIdentityPreserved);
        Run("Native identity parsing rejects malformed and unavailable identities", NativeIdentityParsing);
        Run("Surviving dynamic controls stop refresh before duplicate append", SurvivingControlFailure);
        Run("Detached child-bar cleanup warnings do not prevent a clean root", DetachedBarWarning);
        Run("Native root menu lookup ignores captions", NativeRootLookup);
        Run("Registration uses the resolved native ID, not the proffer return value", NativeRegisteredIdentity);
        Run("A misbound native button is rejected and removed", MisboundButton);
        Run("Icon failure preserves native click routing and is reported", IconFailureDoesNotBreakClick);
        Run("Refreshing a restored native control reapplies its missing icon", RestoreMissingIcon);
        Run("Actual renderer refreshes and restarts with 7 root plus 4 nested scripts, no duplicates, and correct clicks/icons", RebuildAndRestart);
    }

    private static void AddBuiltIns(Fixture fixture)
    {
        fixture.App.Commands.Add("Test.Refresh", BuiltInGroup, RefreshId.ID);
        fixture.App.Commands.Add("Test.OpenFolder", BuiltInGroup, OpenFolderId.ID);
        var refresh = (CommandBarButton)fixture.App.Commands.Item("Test.Refresh", -1).AddControl(fixture.Bar, fixture.Bar.Controls.Count + 1);
        var open = (CommandBarButton)fixture.App.Commands.Item("Test.OpenFolder", -1).AddControl(fixture.Bar, fixture.Bar.Controls.Count + 1);
        refresh.Caption = "Refresh Templates";
        open.Caption = "Open Templates Folder";
    }

    private static void ClearMenu(Fixture fixture, List<Exception> warnings = null)
    {
        QueryTemplateMenuControls.ClearAsync(fixture.Bar, fixture.App.Commands, RefreshId, OpenFolderId,
            ex => { if (warnings != null) warnings.Add(ex); else throw new InvalidOperationException("Unexpected cleanup warning.", ex); },
            () => Task.CompletedTask, CancellationToken.None).GetAwaiter().GetResult();
    }

    private static void UnsupportedTagRegression()
    {
        var fixture = new Fixture();
        AddBuiltIns(fixture);
        const string name = "AxialSqlTools_Template_existing";
        fixture.Registry.RegisterCommand(false, fixture.Handler(name), true);
        var stale = (CommandBarButton)fixture.Bar.Controls[3];
        stale.Tag = name;
        stale.Caption = "existing.sql";
        Equal(string.Empty, stale.Tag);
        // This is the ownership predicate used by the previous cleanup. Native
        // restored buttons retain their command identity but not the Office Tag.
        Equal(0, fixture.Bar.Controls.Count(c => c.Tag.StartsWith("AxialSqlTools_Template_", StringComparison.Ordinal)));
        var unidentifiable = new CommandBarButton { Caption = "another stale script.sql" };
        fixture.Bar.Controls.Add(unidentifiable);
        ClearMenu(fixture);
        Equal(2, fixture.Bar.Controls.Count);
        Assert(stale.Deleted && unidentifiable.Deleted, "Tagless native controls survived cleanup.");
    }

    private static void BuiltInPreflight()
    {
        var fixture = new Fixture();
        AddBuiltIns(fixture);
        fixture.Registry.RegisterCommand(false, fixture.Handler("script"), true);
        var open = fixture.Bar.Controls[2];
        var script = fixture.Bar.Controls[3];
        open.NativeGuid = Guid.NewGuid().ToString(); // Same numeric ID and caption, wrong command group.
        Throws<InvalidOperationException>(() => ClearMenu(fixture));
        Equal(3, fixture.Bar.Controls.Count);
        Assert(!script.Deleted && !open.Deleted, "Preflight failure changed the menu.");
        open.NativeGuid = BuiltInGroup.ToString();
        open.FailCommandInfo = true;
        Throws<InvalidOperationException>(() => ClearMenu(fixture));
        Assert(!script.Deleted, "An unverified built-in action allowed destructive cleanup.");
    }

    private static void BuiltInIdentityTransientFailure()
    {
        var fixture = new Fixture();
        AddBuiltIns(fixture);
        fixture.Registry.RegisterCommand(false, fixture.Handler("script"), true);
        var refresh = fixture.Bar.Controls[1];
        var open = fixture.Bar.Controls[2];
        refresh.CommandInfoCalls = 0;
        open.CommandInfoCalls = 0;
        refresh.CommandInfoSuccessLimit = 1;
        open.CommandInfoSuccessLimit = 1;
        Throws<InvalidOperationException>(() => ClearMenu(fixture));
        Assert(!refresh.Deleted && !open.Deleted, "A second identity-query failure deleted a protected action.");
        Equal(2, fixture.Bar.Controls.Count);
    }

    private static void BuiltInIdentityPreserved()
    {
        var fixture = new Fixture();
        AddBuiltIns(fixture);
        var refresh = fixture.Bar.Controls[1];
        var open = fixture.Bar.Controls[2];
        refresh.Caption = "Localized refresh";
        open.Caption = "Localized folder";
        fixture.Registry.RegisterCommand(false, fixture.Handler("script"), true);
        var duplicate = (CommandBarControl)fixture.App.Commands.Item("script", -1).AddControl(fixture.Bar, 1);
        duplicate.Caption = "Refresh Templates"; // Caption is not command ownership.
        ClearMenu(fixture);
        Equal(2, fixture.Bar.Controls.Count);
        Assert(ReferenceEquals(refresh, fixture.Bar.Controls[1]) && ReferenceEquals(open, fixture.Bar.Controls[2]),
            "Cleanup replaced or removed built-in command objects.");
        Throws<COMException>(() => { var ignored = fixture.Bar.Controls[0]; });
    }

    private static void NativeIdentityParsing()
    {
        var fixture = new Fixture();
        var control = new CommandBarButton { NativeGuid = "invalid", NativeId = 42 };
        Assert(!QueryTemplateMenuControls.TryGetCommandId(fixture.App.Commands, control, out var id) && id == null,
            "Malformed identity was accepted.");
        control.NativeGuid = BuiltInGroup.ToString();
        control.NativeId = 0;
        Assert(!QueryTemplateMenuControls.TryGetCommandId(fixture.App.Commands, control, out id), "Zero native ID was accepted.");
        control.NativeId = 42;
        control.FailCommandInfo = true;
        Assert(!QueryTemplateMenuControls.TryGetCommandId(fixture.App.Commands, control, out id), "Unavailable identity was accepted.");
    }

    private static void SurvivingControlFailure()
    {
        var fixture = new Fixture();
        AddBuiltIns(fixture);
        fixture.Registry.RegisterCommand(false, fixture.Handler("script"), true);
        var script = fixture.Bar.Controls[3];
        script.FailDelete = true;
        var warnings = new List<Exception>();
        Throws<InvalidOperationException>(() => ClearMenu(fixture, warnings));
        Equal(1, warnings.Count);
        Equal(3, fixture.Bar.Controls.Count);
        script.FailDelete = false;
        ClearMenu(fixture);
        Equal(2, fixture.Bar.Controls.Count);
    }

    private static void DetachedBarWarning()
    {
        var fixture = new Fixture();
        AddBuiltIns(fixture);
        var child = fixture.Plugin.AddCommandBarMenu("AxialSqlTools_Folder_warning", MsoBarPosition.msoBarMenuBar, fixture.Bar);
        fixture.Registry.RegisterCommand(false, fixture.Handler("child"), true, child);
        var childButton = child.Controls[1];
        child.FailDelete = true;
        var popup = fixture.Bar.Controls[3];
        var warnings = new List<Exception>();
        ClearMenu(fixture, warnings);
        Equal(1, warnings.Count);
        Equal(2, fixture.Bar.Controls.Count);
        Assert(popup.Deleted && childButton.Deleted, "Child controls or popup survived the detached bar warning.");
    }

    private static void NativeRootLookup()
    {
        var fixture = new Fixture();
        var menuId = new CommandID(BuiltInGroup, 0x1100);
        Throws<InvalidOperationException>(() => fixture.Plugin.GetCommandBar(menuId.Guid, menuId.ID));
        fixture.Bar.Name = "Any localized name";
        fixture.Proffer.BarsByIdentity.Add(menuId, fixture.Bar);
        Assert(ReferenceEquals(fixture.Bar, fixture.Plugin.GetCommandBar(menuId.Guid, menuId.ID)), "Native root identity was not used.");
        Equal(0, fixture.Proffer.BarCalls);
    }

    private static void NativeRegisteredIdentity()
    {
        var fixture = new Fixture();
        fixture.Proffer.ReturnedIdOffset = 1000;
        var handler = fixture.Handler("script");
        fixture.Registry.RegisterCommand(false, handler, true);
        var named = fixture.App.Commands.Item("script", -1);
        var button = (CommandBarButton)fixture.Bar.Controls[1];
        Equal(named.ID, button.NativeId);
        Equal(named.ID, fixture.Command("script").CommandID.ID);
        button.Click();
        Equal(1, handler.Invocations);
    }

    private static void MisboundButton()
    {
        var fixture = new Fixture();
        fixture.App.Commands.ReturnWrongBinding = true;
        Throws<InvalidOperationException>(() => fixture.Registry.RegisterCommand(false, fixture.Handler("script"), true));
        Equal(0, fixture.Bar.Controls.Count);
    }

    private static void IconFailureDoesNotBreakClick()
    {
        var fixture = new Fixture();
        fixture.App.Commands.FailButtonPicture = true;
        var handler = fixture.Handler("script");
        fixture.Registry.RegisterCommand(false, handler, true);
        var button = (CommandBarButton)fixture.Bar.Controls[1];
        button.Click();
        Equal(1, handler.Invocations);
        Equal(1, fixture.Warnings.Count);
        Assert(button.Visible && button.Enabled, "Icon assignment failure disabled the script.");
    }

    private static void RestoreMissingIcon()
    {
        var fixture = new Fixture();
        fixture.Registry.RegisterCommand(false, fixture.Handler("script"), true);
        var button = (CommandBarButton)fixture.Bar.Controls[1];
        button.Picture = null;
        button.Style = MsoButtonStyle.msoButtonAutomatic;
        button.ThrowOnTagWrite = true;
        var replacement = fixture.Handler("script");
        fixture.Registry.RegisterCommand(false, replacement, true);
        Equal(1, fixture.Bar.Controls.Count);
        Assert(button.Picture != null && button.Style == MsoButtonStyle.msoButtonIconAndCaption, "Restored control lost its icon.");
        button.Click();
        Equal(1, replacement.Invocations);
    }

    private static QueryTemplateFolder TestTree()
    {
        string root = Path.Combine(Path.GetTempPath(), "AxialSqlTools-rendering-test");
        var files = Enumerable.Range(0, 7).Select(i => new QueryTemplateFile(Path.Combine(root, "root" + i + ".sql"))).ToList();
        var folders = new List<QueryTemplateFolder>();
        for (int i = 0; i < 4; i++)
        {
            string path = Path.Combine(root, "folder" + i);
            var childFiles = new List<QueryTemplateFile>();
            var childFolders = new List<QueryTemplateFolder>();
            if (i == 3)
            {
                string nested = Path.Combine(path, "nested");
                childFolders.Add(new QueryTemplateFolder(nested, new List<QueryTemplateFolder>(),
                    new List<QueryTemplateFile> { new QueryTemplateFile(Path.Combine(nested, "nested.sql")) }));
            }
            else childFiles.Add(new QueryTemplateFile(Path.Combine(path, "nested.sql")));
            folders.Add(new QueryTemplateFolder(path, childFolders, childFiles));
        }
        return new QueryTemplateFolder(root, folders, files);
    }

    private static Dictionary<string, Handler> PopulateAndClick(Fixture fixture, QueryTemplateFolder tree)
    {
        var handlers = new Dictionary<string, Handler>();
        var warnings = new List<string>();
        var result = QueryTemplateMenuControls.PopulateAsync(tree, fixture.Bar, fixture.Plugin, fixture.Registry,
            file =>
            {
                var handler = fixture.Handler(file.CommandName);
                handlers.Add(file.CommandName, handler);
                return handler;
            }, warnings, () => Task.CompletedTask, CancellationToken.None).GetAwaiter().GetResult();
        Equal(0, result.Failures);
        Equal(11, result.Loaded);
        Equal(0, warnings.Count);
        Equal(13, fixture.Bar.Controls.Count); // Two built-ins, four folders, seven root scripts.
        var buttons = AllScriptButtons(fixture.Bar).ToList();
        Equal(11, buttons.Count);
        foreach (var button in buttons)
        {
            Assert(button.Picture != null && button.Style == MsoButtonStyle.msoButtonIconAndCaption,
                "A root or nested script lost its icon.");
            button.ThrowOnTagWrite = true;
            Equal(string.Empty, button.Tag);
            fixture.App.Commands.CommandInfo(button, out string nativeGuid, out int nativeId);
            var named = fixture.App.Commands.GetByIdentity(nativeGuid, nativeId);
            var expected = handlers[named.Name];
            button.Click();
            Equal(1, expected.Invocations);
        }
        Assert(handlers.Values.All(handler => handler.Invocations == 1), "Native clicks routed to the wrong script.");
        Equal(13, fixture.App.Commands.Count); // Persisted named commands cannot accumulate on refresh.
        Equal(11, fixture.MenuService.Count);
        return handlers;
    }

    private static IEnumerable<CommandBarButton> AllScriptButtons(CommandBar bar)
    {
        foreach (CommandBarControl control in bar.Controls)
        {
            if (control is CommandBarPopup popup)
            {
                foreach (var child in AllScriptButtons(popup.CommandBar)) yield return child;
            }
            else if (control is CommandBarButton button && button.NativeGuid != BuiltInGroup.ToString("B"))
                yield return button;
        }
    }

    private static void RebuildAndRestart()
    {
        var fixture = new Fixture();
        AddBuiltIns(fixture);
        var tree = TestTree();
        var previousHandlers = PopulateAndClick(fixture, tree);
        for (int refresh = 0; refresh < 2; refresh++)
        {
            var previousButtons = AllScriptButtons(fixture.Bar).ToList();
            ClearMenu(fixture);
            Assert(previousButtons.All(button => button.Deleted), "Refresh left an old root or nested button behind.");
            RemoveUnused(fixture, new HashSet<string>(previousHandlers.Keys));
            var replacement = PopulateAndClick(fixture, tree);
            Assert(previousHandlers.Values.All(handler => handler.Invocations == 1), "Refresh invoked a stale handler.");
            previousHandlers = replacement;
        }
        var persistedButtons = AllScriptButtons(fixture.Bar).ToList();
        var restarted = new Fixture(fixture.App, fixture.Bar, fixture.Group);
        Equal(0, restarted.MenuService.Count);
        ClearMenu(restarted);
        Assert(persistedButtons.All(button => button.Deleted), "A fresh registry failed to clean persisted native buttons.");
        RemoveUnused(restarted, new HashSet<string>(previousHandlers.Keys));
        PopulateAndClick(restarted, tree);
        Assert(previousHandlers.Values.All(handler => handler.Invocations == 1), "Restart reused a previous session callback.");
    }
}
