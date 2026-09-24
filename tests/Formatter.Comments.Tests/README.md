# Comment-preserving formatter regressions

Run from the repository root with the .NET 8 SDK:

```sh
dotnet run --project tests/Formatter.Comments.Tests
```

The harness compiles the production `TsqlFormatter`, `TsqlFormatterCommentInterleaver`, and `TsqlSelectFieldIndentation` files directly. Only the settings container and SSMS logger are stubbed; a formatter error fails the test instead of being logged and ignored. No SQL Server connection is used.

The main extension continues using the ScriptDOM assembly installed with SSMS. The NuGet dependency below belongs only to this standalone test project.

Compatibility checks:

```sh
dotnet run --project tests/Formatter.Comments.Tests -p:ScriptDomVersion=170.157.0
dotnet run --project tests/Formatter.Comments.Tests -p:ScriptDomVersion=170.191.0
dotnet run --project tests/Formatter.Comments.Tests -p:ScriptDomVersion=180.102.0
```

Validated results:

| ScriptDOM version | Passed | Failed |
| --- | ---: | ---: |
| 170.157.0 | 146 | 0 |
| 170.191.0 | 150 | 0 |
| 180.102.0 | 150 | 0 |

The older build lacks the native `PreserveComments` option, so four native-preference combinations do not apply.

Coverage includes:

- The Settings preview example, including inline comments after BEGIN and trailing multiline comments.
- Both Disregard modes, with all Axial transforms disabled/enabled.
- Exact comment text, order, and count, including strings containing comment markers, nested block comments, Unicode, CRLF/LF/CR input, and comment-only scripts.
- JOIN, CASE, SELECT lists, procedure declarations/calls, DECLARE lists, END, semicolons, commas, and GO boundaries.
- Equivalent comment-free SQL before/after preservation, valid parsing, and stable output on a second formatting pass.
- Axial/native comment preferences, avoiding generator option leakage, and avoiding duplicate native comments in the direct helper.
- Linear-memory alignment against a dynamic-programming LCS reference on 1,000 deterministic random/repeated-token sequences.
- A 2,000-statement script with inserted aliases and 2,000 comments. Formatting allocates approximately 91-92 MiB in the tested builds; a 256 MiB allocation ceiling guards against restoring the old quadratic table. Timing is reported without a brittle timing assertion.

The formatter now restores comments after all code formatting passes. It anchors inline comments to preceding code and standalone comments to following code, without copying source indentation into formatted statements. A trailing line comment is placed after an adjacent comma/semicolon so it cannot swallow the delimiter. Comment contents, including whitespace inside multiline comments, remain unchanged.

The final result must parse and retain every generated non-comment SQL token and every original comment token. If preservation cannot satisfy those checks, formatting throws before the editor command replaces the user's SQL.

Windows follow-up: build/install the VSIX, check the Settings preview with Preserve comments and Disregard toggled, then format the example in an SSMS query editor. This harness does not exercise the SSMS shell or WPF UI.
