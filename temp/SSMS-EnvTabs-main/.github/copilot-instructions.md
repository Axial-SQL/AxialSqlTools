# SSMS EnvTabs - AI Coding Instructions

## Overview
SSMS EnvTabs is a Visual Studio Extension (VSIX) for SQL Server Management Studio (SSMS) 2017+ (shell-based). It groups and colors query tabs based on connection properties (Server/Database) defined in user configuration.

## Architecture & Key Components
- **Entry Point**: `SSMS_EnvTabsPackage.cs` initializes the extension as an `AsyncPackage`.
- **Event Loop**: `RdtEventManager.cs` subscribes to `IVsRunningDocTableEvents` to detect when query windows are opened, closed, or switched. 
  - **Connection Changes**: Uses `OnAfterAttributeChange` events combined with a polling retry mechanism (`ScheduleRenameRetry`) to handle SSMS's delayed UI updates when switching connections.
- **Logic**:
  - `TabRuleMatcher.cs`: Matches current connection info against loaded rules.
  - `TabRenamer.cs`: Renames window captions (e.g., "Prod-1") using `IVsWindowFrame.SetProperty`.
  - `TabGroupColorSolver.cs`: Implements SSMS's hashing algorithm to calculate salts required to target specific colors.
  - `ColorByRegexConfigWriter.cs`: Writes regex-based coloring rules to a temp file (`ColorByRegexConfig.txt`) consumed by SSMS/MIDS. It uses the Solver to automatically inject salts.
- **Configuration**: User rules are loaded from `%USERPROFILE%\Documents\SSMS EnvTabs\TabGroupConfig.json`.

### Coloring Logic & Challenges
- **Mechanism**: The extension writes a `ColorByRegexConfig.txt` file which SSMS reads.
- **Regex Format**: We generate regexes based on **Filenames only** (e.g., `(query.sql|script.sql)`) to avoid path dependency issues.
- **Color Assignment Algorithm** (Reverse Engineered):
  - SSMS calculates a hash using **.NET Framework Legacy 32-bit String.GetHashCode**.
  - The Color Index (0-15) is determined by `Math.Abs(Hash) % 16`.
- **Automated Salting**:
  - Users specify a `ColorIndex` (0-15) in their config.
  - The extension (via `TabGroupColorSolver`) bruteforces a short suffix string (salt) such that `Hash(Regex + (?#salt:xyz)) % 16 == TargetIndex`.
  - This salt is appended to the regex as a comment, invisible to matching logic but affecting the hash.
- **Color Persistence**: SSMS reads `ColorByRegexConfig.txt` when it changes. We force updates when tabs open or close.

## Developer Workflows

### Debugging
- **Start Action**: The project defaults to `devenv.exe /rootsuffix Exp`. To debug in SSMS, change project properties → Debug → Start external program to your `ssms.exe` path (e.g., `C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\Ssms.exe`).
- **Command Arguments**: Use `/log` to enable SSMS logging if needed.
- **Attaching**: You can attach to a running `Ssms.exe` process if the extension is installed.

### Logging
- **Runtime Log**: The extension writes its own log to `%LocalAppData%\SSMS EnvTabs\runtime.log` via `EnvTabsLog.cs`. Check this first for logic errors. **Note: This file logging is for development only and will be removed in the final release.**
- **ActivityLog**: Standard VS activity log is at `%AppData%\Microsoft\SQL Server Management Studio\...\ActivityLog.xml`.

### Build & Deploy
- **Build**: Uses standard MSBuild with VS SDK targets. **Never attempt to build the project using tools. The user will handle building and installing via Visual Studio.**
- **Install**: Use `VSIXInstaller.exe` manually or via script (see `dev\references.md`) to install the `.vsix` into SSMS.

## Project Conventions

### Threading
- **Strict UI Thread Usage**: Most VS Shell interactions (RDT, Window Frames) must happen on the UI thread.
- **Pattern**: Use `await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync()` or `ThreadHelper.ThrowIfNotOnUIThread()` before accessing `IVs*` services.

### Configuration Pattern
- **External Config**: Unlike typical VS extensions using `DialogPage`, this project uses a standalone JSON file in the User's Documents folder to allow sharing/scripting of rules.
- **Reloading**: Config is reloaded automatically or on specific triggers (check `RdtEventManager.cs` for logic).

### VS Interop
- **Services**: `SVsRunningDocumentTable`, `SVsUIShellOpenDocument` are key services.
- **Properties**: Captions are modified via `VSFPROPID_OwnerCaption` or fallback to `VSFPROPID_Caption`.
