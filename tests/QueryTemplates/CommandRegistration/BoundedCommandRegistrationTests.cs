using Microsoft.VisualStudio.CommandBars;
using System;
using System.Collections.Generic;

internal static partial class CommandRegistrationTests
{
    private static void RunBoundedRegistrationTests()
    {
        Run("Empty registry cleanup performs no global command enumeration or lookup", EmptyRegistryCleanupIsBounded);
        Run("Cleanup looks up only obsolete commands tracked by this registry", TrackedCleanupIsBounded);
        Run("An already missing tracked command still loses its old callback", AlreadyMissingTrackedCommand);
        Run("Unexpected lookup failures are reported and obsolete callbacks are detached", LookupFailureDetachesCallback);
        Run("Cleanup never deletes a tracked name whose GUID or ID changed", ChangedIdentityIsNotDeleted);
        Run("Fresh-menu registration performs one native identity query per script", FreshMenuIdentityQueriesAreLinear);
        Run("Append-only registration still rejects incorrectly bound controls", AppendOnlyValidatesNativeBinding);
    }

    private static void EmptyRegistryCleanupIsBounded()
    {
        var fixture = new Fixture();
        for (int i = 0; i < 1000; i++)
            fixture.App.Commands.Add("Unrelated.Command" + i, Guid.NewGuid(), 10000 + i);
        RemoveUnused(fixture, new HashSet<string>());
        Equal(0, fixture.App.Commands.EnumerationCalls);
        Equal(0, fixture.App.Commands.ItemCalls);
        Equal(0, fixture.App.Commands.DeletedCount);
        Equal(1000, fixture.App.Commands.Count);
    }

    private static void TrackedCleanupIsBounded()
    {
        var fixture = new Fixture();
        for (int i = 0; i < 1000; i++)
            fixture.App.Commands.Add("Unrelated.Command" + i, Guid.NewGuid(), 10000 + i);
        fixture.Registry.RegisterCommand(false, fixture.Handler("AxialSqlTools_Template_removed"), true, appendOnly: true);
        fixture.Registry.RegisterCommand(false, fixture.Handler("AxialSqlTools_Template_retained"), true, appendOnly: true);
        fixture.App.Commands.ItemCalls = 0;
        RemoveUnused(fixture, new HashSet<string> { "AxialSqlTools_Template_retained" });
        Equal(1, fixture.App.Commands.ItemCalls);
        Equal(0, fixture.App.Commands.EnumerationCalls);
        Equal(1, fixture.App.Commands.DeletedCount);
        Equal(1001, fixture.App.Commands.Count);
        Equal(1, fixture.MenuService.Count);
    }

    private static void AlreadyMissingTrackedCommand()
    {
        var fixture = new Fixture();
        const string name = "AxialSqlTools_Template_removed";
        var handler = fixture.Handler(name);
        fixture.Registry.RegisterCommand(false, handler, true);
        var stale = fixture.Command(name);
        fixture.App.Commands.Item(name, -1).Delete();
        RemoveUnused(fixture, new HashSet<string>());
        Equal(0, fixture.MenuService.Count);
        Equal(0, fixture.App.Commands.EnumerationCalls);
        stale.QueryStatus();
        Assert(!stale.Enabled && !stale.Visible && !stale.Supported, "Missing native command retained an enabled callback.");
        stale.Invoke();
        Equal(0, handler.Invocations);
    }

    private static void LookupFailureDetachesCallback()
    {
        var fixture = new Fixture();
        const string name = "AxialSqlTools_Template_removed";
        var handler = fixture.Handler(name);
        fixture.Registry.RegisterCommand(false, handler, true);
        var stale = fixture.Command(name);
        fixture.App.Commands.FailLookupName = name;
        Throws<InvalidOperationException>(() => RemoveUnused(fixture, new HashSet<string>()));
        Equal(0, fixture.MenuService.Count);
        Equal(1, fixture.App.Commands.Count);
        Equal(0, fixture.App.Commands.DeletedCount);
        Equal(0, fixture.App.Commands.EnumerationCalls);
        stale.Invoke();
        Equal(0, handler.Invocations);
    }

    private static void ChangedIdentityIsNotDeleted()
    {
        for (int variation = 0; variation < 2; variation++)
        {
            var fixture = new Fixture();
            const string name = "AxialSqlTools_Template_reused_name";
            var handler = fixture.Handler(name);
            fixture.Registry.RegisterCommand(false, handler, true);
            var stale = fixture.Command(name);
            var native = fixture.App.Commands.Item(name, -1);
            if (variation == 0) native.Guid = Guid.NewGuid().ToString();
            else native.ID += 1000;
            Throws<InvalidOperationException>(() => RemoveUnused(fixture, new HashSet<string>()));
            Equal(0, fixture.App.Commands.DeletedCount);
            Equal(1, fixture.App.Commands.Count);
            Equal(0, fixture.MenuService.Count);
            stale.Invoke();
            Equal(0, handler.Invocations);
        }
    }

    private static void FreshMenuIdentityQueriesAreLinear()
    {
        var fixture = new Fixture();
        AddBuiltIns(fixture);
        const int scriptCount = 512;
        var handlers = new List<Handler>();
        fixture.App.Commands.CommandInfoCallCount = 0;
        for (int i = 0; i < scriptCount; i++)
        {
            var handler = fixture.Handler("AxialSqlTools_Template_fresh_" + i);
            handlers.Add(handler);
            fixture.Registry.RegisterCommand(false, handler, true, fixture.Bar, appendOnly: true);
        }
        Equal(scriptCount, fixture.App.Commands.CommandInfoCallCount);
        Equal(scriptCount + 2, fixture.Bar.Controls.Count);
        Equal(0, fixture.App.Commands.EnumerationCalls);
        ((CommandBarButton)fixture.Bar.Controls[3]).Click();
        ((CommandBarButton)fixture.Bar.Controls[scriptCount + 2]).Click();
        Equal(1, handlers[0].Invocations);
        Equal(1, handlers[scriptCount - 1].Invocations);
    }

    private static void AppendOnlyValidatesNativeBinding()
    {
        var fixture = new Fixture();
        fixture.App.Commands.ReturnWrongBinding = true;
        Throws<InvalidOperationException>(() => fixture.Registry.RegisterCommand(false,
            fixture.Handler("AxialSqlTools_Template_misbound"), true, appendOnly: true));
        Equal(0, fixture.Bar.Controls.Count);
        Equal(1, fixture.App.Commands.CommandInfoCallCount);
    }
}
