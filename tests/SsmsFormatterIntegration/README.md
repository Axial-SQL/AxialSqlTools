# SSMS formatter integration

Run the focused checks with .NET 8:

```sh
dotnet run --project tests/SsmsFormatterIntegration/SsmsFormatterIntegration.Tests.csproj
```

The tests link the production reflection bridge and formatting pipeline. A handwritten test double exposes the SSMS internal API shapes, and real ScriptDOM parsers/generators produce the SQL. These checks cover settings refresh, document-buffer forwarding, parser/generator selection, casing/alignment/indentation, all four native/Axial comment-setting combinations, individual and combined Axial options, parse errors, failed settings loading, cancellation, and missing API contracts. They do not execute the installed SSMS formatter DLL or validate shell service discovery.

## Integration and precedence

The Format command uses its original DTE `ActiveDocument`, `TextSelection`, and `TextDocument` path for reading and replacing SQL. `FormatSettingsLoader.LoadAsync` receives `SqlFormatterExtension.ExtensibilityInstance` with null view/buffer arguments and loads SSMS global settings. Document-specific `.editorconfig` overrides are not applied. No ComponentModelHost or editor adapter assembly lookup is needed. The bridge invokes `SqlFormatHelper.ToScriptGeneratorOptions`, `CreateParser`, and `CreateScriptGenerator`, including their non-public members, and passes the resulting ScriptDOM objects into Axial's existing post-processing pipeline.

- SSMS controls the initial SQL version, engine, layout, casing, and comments.
- Enabled Axial options, including Shift+Format overrides, run afterward and may change that layout. Existing Axial transforms that explicitly insert spaces retain that behavior even when SSMS uses tabs.
- Axial's Preserve comments option forces preservation. When unchecked, the native SSMS comment setting applies. The interleaver is only used when a generator cannot preserve comments itself.
- The Settings preview uses SSMS global settings without a document's `.editorconfig`.
- Scripting/export callers retain the standalone formatter entry point, so formatting generated scripts does not acquire settings from an unrelated active query.
- No formatter DLL, decompiled SSMS implementation, editor SDK package, or new settings storage is distributed.

The bridge targets the internal API in the supplied decompilation of `Microsoft.SqlServer.Management.SqlFormatter` version `22.200.0.0`. A missing/uninitialized formatter, failed settings load, or incompatible API/ScriptDOM assembly produces an error before the Format command edits SQL. There is no silent fallback to another formatting style.

## Windows / SSMS validation

1. Build and install the VSIX in the SSMS version containing the above formatter. Open a SQL query, then invoke Axial Format before invoking the native formatter, to verify normal extension activation.
2. With Axial options unchecked, compare native SSMS Format and Axial Format for keyword case, indentation, clause layout, comments, SQL version, and engine settings. The Axial command retains its existing surrounding-whitespace handling.
3. Change SSMS SQL Formatter settings and format again without restarting. Check the Settings preview after toggling an Axial option.
4. Format both a saved SQL file and an unsaved query using the original DTE path. Axial uses SSMS global settings in both cases, without document-specific `.editorconfig` overrides.
5. Test whole-document and selected-text formatting, Shift+Format overrides/cancel, and enabled Axial transformations together. Check comments appear exactly once.
6. Open Settings without an active SQL query. If SSMS's formatter has not initialized, preview should show an explanatory message while the Settings UI remains usable. Open a query and toggle an option to retry.
7. Verify a parse error or unavailable/incompatible formatter leaves query text unchanged. Switch query windows while settings load and confirm formatting stops if the active DTE document changes.
8. Smoke-test scripting/export features, which still use their standalone path.

