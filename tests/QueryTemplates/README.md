# Query Templates regression checks

These standalone checks require the .NET 10 SDK and do not require SSMS or NuGet test packages. They compile the production sources with C# 7.3.

From the repository root:

```sh
dotnet run --project tests/QueryTemplates/QueryTemplateCatalogTests.csproj
dotnet run --project tests/QueryTemplates/CommandRegistration/CommandRegistrationTests.csproj
dotnet run --project tests/QueryTemplates/RefreshScheduling/QueryTemplateRefreshCoordinatorTests.csproj
dotnet run --project tests/QueryTemplates/PackageOrchestration/PackageOrchestrationTests.csproj
```

The catalog checks cover sorting, root and nested files, all file extensions, stable identities, additions/renames/deletions, unavailable folders, links, and cancellation. Registration and menu checks use a stub host with persistent native GUID/ID bindings, empty tags, and one-based control positions. They exercise the production cleanup and menu population code, including root and nested button clicks, icon restoration, repeated refreshes, simulated restarts, failed bindings, and protection of the two built-in actions. The host rejects global command enumeration, and a large-menu check verifies linear native identity queries when adding scripts. Cancellation and cyclic popup checks exercise the actual asynchronous menu code.

Performance checks run the production renderer on both cold and warm command registries. For 512 sibling folders, normal attachment requires 512 parent index probes. For 512 scripts under a fixed clock and ready shell, population uses 15 batch waits, 512 binding checks, and 1,024 cold or 512 warm command lookups. These are operation counts, not host timing measurements. Unexpected popup attachment and returned-bar ownership are tested before any failed-creation cleanup can delete a native bar.

Scheduling checks exercise the production refresh coordinator, UI gate, and batcher: simultaneous and reentrant requests share one run, completed/failed runs can be retried explicitly, duplicate scheduler callbacks cannot repeat work, and modal waits yield and honor cancellation. The gate is checked again after background discovery. Native work is grouped into short batches with a 16 ms target and a maximum of 32 checkpoints, while cancellation and modal state are checked per checkpoint. A single native call can exceed the time target; the next checkpoint then yields. Deterministic tests use a fake clock to check both limits and exclude time spent waiting from the next batch's budget. These tests do not simulate the complete SSMS shell or its COM message pump.

Package orchestration checks compile the actual templates package partial and rendering code. The test host pumps only normal-priority UI work, withholds all idle work, and never opens a menu. Startup must still create the complete root/nested tree and make it usable. The host also withholds the deferred global command-status invalidation, so the test verifies direct control state rather than relying on it. Modal-startup and cancellation checks prevent menu mutation before the shell is ready or after disposal. This models the scheduling contract; it does not reproduce the complete native SSMS host.

## SSMS 22 validation

Build and install the VSIX using the manual build workflow or a local Windows build, then verify:

1. Install over the previous build with its existing menu customizations intact. Configure a folder with seven root scripts and four nested folders. Start SSMS repeatedly and verify one copy of every folder/script, icons on root and nested scripts, and no dead entries. The first successful refresh should remove old duplicated entries automatically.
2. Select a root script and a nested script. Verify normal insertion and Shift-click opening a new query.
3. Add, rename, and remove files and folders, then select Refresh Templates repeatedly. Check ordering, labels, inserted content, and absence of duplicates.
4. Temporarily make the root folder unavailable and refresh. The existing menu should remain, with a failure message. Restore access and refresh without restarting SSMS.
5. Make one nested folder inaccessible. Available root/sibling scripts should still load, with a warning in the status bar and details in the extension log.
6. Open Templates Folder and confirm it opens Explorer without rebuilding the menu. Refresh success should appear in the SSMS status bar rather than a completion dialog.
7. Start SSMS with the Connect dialog open. Leave it open, type in it, move it, and cancel or connect. The dialog must stay responsive; template menu mutation should wait until the dialog closes. Repeat with a large templates folder and with an empty folder.
8. Open a modal dialog while a large refresh is running. Menu work must pause and resume after it closes. Click Refresh repeatedly while loading and verify one rebuild, no duplicate controls, and responsive input. Close SSMS during a refresh and verify shutdown completes.
9. Compare first startup and subsequent Refresh with the same large folder, including many sibling subfolders. Keep the dropdown open for one run. Check that folders become usable promptly, root and nested scripts retain their icons and click behavior, and typing/moving windows stays responsive during loading. Use extension log timings to compare native menu work with time spent waiting for UI access.
10. Start SSMS without touching Query Templates. Close Connect, wait for the loaded status/log message, and only then open the menu. Every populated root/nested folder should already be usable. Repeat with the SQL Server builds download unavailable; template loading and package initialization must not wait for that download.

The extension log records discovery, old-control cleanup, completion, and elapsed time. It also records menu-work duration, batch wait count, and total batch wait time to distinguish scheduler delays from native creation cost. If a hang persists, record the last query-template log stage and capture the SSMS process stacks while it is unresponsive. Passing the standalone checks is not proof of live SSMS startup behavior or a measured SSMS speedup.

Logs are under `%LOCALAPPDATA%\AxialSQL\AxialSQLToolsLog\log_YYYY-MM-DD.log`. Package initialization, template startup request, and normal-priority worker start are separate entries. If package initialization itself only starts when a menu is opened, that is host package activation, before the template worker runs. The extension retains supported background auto-loading; it cannot force SSMS to activate the package before its host policy permits. Once activated, template creation uses normal-priority dispatcher work and timed batch pauses rather than shell idle notifications. Modal dialogs still deliberately postpone native menu mutation.

No workflow is triggered automatically by these checks.
