# Query Templates regression checks

These standalone checks require the .NET 10 SDK and do not require SSMS or NuGet test packages. They compile the production sources with C# 7.3.

From the repository root:

```sh
dotnet run --project tests/QueryTemplates/QueryTemplateCatalogTests.csproj
dotnet run --project tests/QueryTemplates/CommandRegistration/CommandRegistrationTests.csproj
```

The catalog checks cover sorting, root and nested files, all file extensions, stable identities, additions/renames/deletions, unavailable folders, links, and cancellation. Registration and menu checks use a stub host with persistent native GUID/ID bindings, empty tags, and one-based control positions. They exercise the production cleanup and menu population code, including root and nested button clicks, icon restoration, repeated refreshes, simulated restarts, failed bindings, and protection of the two built-in actions. They do not simulate the complete SSMS shell.

## SSMS 22 validation

Build and install the VSIX using the manual build workflow or a local Windows build, then verify:

1. Install over the previous build with its existing menu customizations intact. Configure a folder with seven root scripts and four nested folders. Start SSMS repeatedly and verify one copy of every folder/script, icons on root and nested scripts, and no dead entries. The first successful refresh should remove old duplicated entries automatically.
2. Select a root script and a nested script. Verify normal insertion and Shift-click opening a new query.
3. Add, rename, and remove files and folders, then select Refresh Templates repeatedly. Check ordering, labels, inserted content, and absence of duplicates.
4. Temporarily make the root folder unavailable and refresh. The existing menu should remain, with a failure message. Restore access and refresh without restarting SSMS.
5. Make one nested folder inaccessible. Available root/sibling scripts should still load, with a warning in the status bar and details in the extension log.
6. Open Templates Folder and confirm it opens Explorer without rebuilding the menu. Refresh success should appear in the SSMS status bar rather than a completion dialog.

No workflow is triggered automatically by these checks.
