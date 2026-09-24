# Theme validation

Run the source audit with Python 3:

```sh
python tests/Themes/validate_themes.py
```

Compile the production WPF markup and theme code with the .NET 8 SDK and .NET Framework 4.8 reference assemblies:

```sh
python tests/Themes/validate_themes.py --build
```

The harness includes all 27 production XAML files, the shared theme controller, generated pivot styles, and the actual chart and SQL-preview theme methods. It uses the production AvalonEdit/OxyPlot versions and Visual Studio SDK 17.14.40264. Unrelated SSMS/database event handlers are stubbed. NuGet restores the reference assemblies; SQL Server and SSMS are not required for this compilation check. `--dotnet` accepts an explicit SDK executable path.

The audit checks all 26 views for shared resources, themed foreground/background pairs, and live theme subscriptions. It also rejects missing brush resources, duplicate shared keys, fixed colors in generated styles, and fixed UI colors outside the shared design-time palette. The connection-color preview is intentionally exempt because it displays the user's chosen color.

## SSMS visual check (Windows required)

Build and install the extension in SSMS. Check Light, Dark, Blue (where available), and Windows High Contrast. Switch themes with populated tool windows open, including hidden tabs, then reopen closed dialogs.

- Open Settings, About, Formatter Options, SSMS Formatter Settings, Query History, Snippet Manager, Quick Search, Data Import/Transfer/Compare, saved connections, object selection, export confirmations, Statistics Summary, SQL Server Builds, Pivot Grid/details, Grid to Email, Sync to GitHub, health dashboards, and the query-safety warning.
- Check hover, pressed, disabled, keyboard focus, selected/unselected text, tabs, scrollbars, password fields, tooltips, context menus, and editable/read-only combo boxes. In Data Compare, verify that table names remain visible and selectable.
- Open Query History calendars. Check empty/populated dates, typing, month/year navigation, keyboard selection, and disabled dates.
- Populate a pivot, select cells, open details, sort, and scroll to its pinned total. Switch themes and check all cells, headers, totals, and selection text without rebuilding the pivot.
- Check Quick Search SQL syntax, selected text, and existing match highlights. Check loaded health charts and legends without waiting for a data refresh. Repeated theme switches must not progressively alter their colors.
- Verify selected SQL Server Builds rows and links, and selected connection-color rules in Settings. User-specified connection colors must remain unchanged.
- Close a dialog or tool window during a theme change. Reopen it and verify the current theme without stale subscriptions or exceptions.

The compilation check does not render WPF, run the full extension, or validate VSIX packaging. Windows/SSMS must validate those behaviors. Native operating-system dialogs, such as file and color pickers, retain their host behavior.
