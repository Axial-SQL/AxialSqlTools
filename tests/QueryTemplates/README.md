# Query Templates regression checks

These standalone checks require the .NET 10 SDK and do not require SSMS or NuGet test packages. They compile the production sources with C# 7.3.

From the repository root:

```sh
dotnet run --project tests/QueryTemplates/QueryTemplateCatalogTests.csproj
dotnet run --project tests/QueryTemplates/CommandRegistration/CommandRegistrationTests.csproj
dotnet run --project tests/QueryTemplates/RefreshScheduling/QueryTemplateRefreshCoordinatorTests.csproj
```

The catalog checks cover sorting, root and nested files, all file extensions, stable identities, additions/renames/deletions, unavailable folders, links, and cancellation. Registration and menu checks use a stub host with persistent native GUID/ID bindings, empty tags, and one-based control positions. They exercise the production cleanup and menu population code, including root and nested button clicks, icon restoration, repeated refreshes, simulated restarts, failed bindings, and protection of the two built-in actions. The host rejects global command enumeration, and a large-menu check verifies linear native identity queries when adding scripts. Cancellation and cyclic popup checks exercise the actual asynchronous menu code.

Scheduling checks exercise the production refresh coordinator and UI gate: simultaneous and reentrant requests share one run, completed/failed runs can be retried explicitly, duplicate scheduler callbacks cannot repeat work, and modal waits yield and honor cancellation. The gate is checked again after background discovery. These tests do not simulate the complete SSMS shell or its COM message pump.

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

The extension log records discovery, old-control cleanup, completion, and elapsed time. If a hang persists, record the last query-template log stage and capture the SSMS process stacks while it is unresponsive. Passing the standalone checks is not proof of live SSMS startup behavior.

No workflow is triggered automatically by these checks.
