using AxialSqlTools;
using Microsoft.SqlServer.Management.SqlFormatter;
using Microsoft.SqlServer.TransactSql.ScriptDom;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

internal static class Program
{
    private static int assertions;
    private static Task<SsmsFormatterContext> Create(object buffer = null, CancellationToken token = default, bool disregardSsmsSettings = false)
        => SsmsFormatterReflection.CreateAsync(Assembly.GetExecutingAssembly(), buffer, token, disregardSsmsSettings);
    private static void Check(bool condition, string message)
    {
        assertions++;
        if (!condition) throw new Exception(message);
    }
    private static async Task Reject(Func<Task> action, string expected)
    {
        try { await action(); }
        catch (Exception ex)
        {
            Check(!(ex is TargetInvocationException) && ex.Message.Contains(expected), ex.ToString());
            return;
        }
        throw new Exception("Expected failure: " + expected);
    }
    private static string Format(SsmsFormatterContext context, string sql, SettingsManager.TSqlCodeFormatSettings settings = null)
        => TSqlFormatter.FormatCode(sql, settings ?? new SettingsManager.TSqlCodeFormatSettings(), context.Parser, context.Generator);
    private static void Valid(string sql, SsmsFormatterContext context)
    {
        using var reader = new StringReader(sql);
        context.Parser.Parse(reader, out var errors);
        Check(errors.Count == 0, "Invalid formatted SQL: " + string.Join("; ", errors.Select(e => e.Message)));
    }
    private static int Comments(string sql)
    {
        using var reader = new StringReader(sql);
        var tokens = new TSql170Parser(true).GetTokenStream(reader, out var errors);
        return tokens.Count(t => t.TokenType == TSqlTokenType.SingleLineComment || t.TokenType == TSqlTokenType.MultilineComment);
    }
    private static async Task Main()
    {
        var context = await Create();
        Check(context.Parser is TSql160Parser && context.Generator is Sql160ScriptGenerator, "Use SSMS factories/version.");
        Check(context.Generator.Options.AlignClauseBodies, "Keep SSMS alignment.");
        Check(context.Generator.Options.IndentationSize == 2, "Use global indentation.");
        var lower = Format(context, "SELECT [value] FROM dbo.t;");
        Check(lower.Contains("select") && lower.Contains("from"), "Use SSMS keyword casing.");
        Valid(Format(context, "SELECT \"quoted identifier\" FROM dbo.t;"), context);
        var buffer = new object();
        var document = await Create(buffer);
        Check(ReferenceEquals(FormatSettingsLoader.LastBuffer, buffer), "Forward active document buffer.");
        Check(document.Generator.Options.IndentationSize == 6, "Use document overrides.");
        Check(!new SettingsManager.TSqlCodeFormatSettings().disregardSsmsFormatterSettings, "Use SSMS settings by default.");
        Check(!new SettingsManager.TSqlCodeFormatSettings { disregardSsmsFormatterSettings = true }.HasAnyFormattingEnabled(), "Formatter mode is not an Axial post-processing option.");
        var mappingsBefore = SqlFormatHelper.OptionsMappings;
        var defaults = await Create(buffer, disregardSsmsSettings: true);
        Check(SqlFormatHelper.OptionsMappings == mappingsBefore + 1, "Disregard mode still applies SSMS settings before creating native instances.");
        Check(!ReferenceEquals(defaults.Parser, SqlFormatHelper.LastParser), "Replace the SSMS parser with a fresh instance.");
        Check(!ReferenceEquals(defaults.Generator, SqlFormatHelper.LastGenerator), "Replace the SSMS generator with a fresh instance.");
        Check(SqlFormatHelper.LastParser.QuotedIdentifier && !defaults.Parser.QuotedIdentifier, "Use the original parser constructor argument, false.");
        Check(SqlFormatHelper.LastGenerator.Options.KeywordCasing == KeywordCasing.Lowercase
            && SqlFormatHelper.LastGenerator.Options.IndentationSize == 6, "Create the native generator with SSMS settings before replacing it.");
        Check(defaults.Parser.GetType() == document.Parser.GetType(), "Default parser retains effective SQL version.");
        Check(defaults.Generator.GetType() == document.Generator.GetType(), "Default generator retains effective SQL version.");
        var expectedGenerator = new Sql160ScriptGenerator();
        var expectedOptions = expectedGenerator.Options;
        foreach (var property in typeof(SqlScriptGeneratorOptions).GetProperties().Where(p => p.CanRead && p.GetIndexParameters().Length == 0))
            Check(Equals(property.GetValue(defaults.Generator.Options), property.GetValue(expectedOptions)), "Default option: " + property.Name);
        using (var reader = new StringReader("SELECT a, b FROM dbo.t;"))
        {
            var fragment = new TSql160Parser(false).Parse(reader, out var errors);
            expectedGenerator.GenerateScript(fragment, out var expected);
            Check(Format(defaults, "SELECT a, b FROM dbo.t;") == expected, "Default mode matches a freshly created generator.");
        }
        Check(document.Generator.Options.KeywordCasing == KeywordCasing.Lowercase && document.Generator.Options.IndentationSize == 6, "Default mode does not mutate SSMS options.");
        var resumed = await Create(buffer);
        Check(resumed.Generator.Options.KeywordCasing == KeywordCasing.Lowercase && resumed.Generator.Options.IndentationSize == 6, "Switching back restores SSMS options.");
        var withAxialOptions = Format(defaults, "SELECT DISTINCT TOP 5 a, b FROM dbo.t;", new SettingsManager.TSqlCodeFormatSettings
        {
            disregardSsmsFormatterSettings = true,
            breakSelectFieldsAfterTopAndUnindent = true
        });
        Check(withAxialOptions.Replace("\r", "").Contains("\n" + new string(' ', expectedOptions.IndentationSize) + "a"), "Axial transforms still run with default formatting.");
        Valid(withAxialOptions, defaults);
        SqlFormatHelper.Use170Implementation = true;
        var newerDefaults = await Create(buffer, disregardSsmsSettings: true);
        Check(newerDefaults.Parser is TSql170Parser && newerDefaults.Generator is Sql170ScriptGenerator,
            "Use the actual SSMS concrete types, not the settings enum.");
        Check(!ReferenceEquals(newerDefaults.Parser, SqlFormatHelper.LastParser)
            && !ReferenceEquals(newerDefaults.Generator, SqlFormatHelper.LastGenerator), "Create fresh instances for SQL 170 too.");
        var expected170 = new Sql170ScriptGenerator();
        foreach (var property in typeof(SqlScriptGeneratorOptions).GetProperties().Where(p => p.CanRead && p.GetIndexParameters().Length == 0))
            Check(Equals(property.GetValue(newerDefaults.Generator.Options), property.GetValue(expected170.Options)), "SQL 170 constructor default: " + property.Name);
        SqlFormatHelper.Use170Implementation = false;
        var normal = await Create(buffer);
        Check(ReferenceEquals(normal.Parser, SqlFormatHelper.LastParser) && ReferenceEquals(normal.Generator, SqlFormatHelper.LastGenerator),
            "Normal mode retains the configured SSMS instances.");
        FormatSettingsLoader.Casing = KeywordCasing.Uppercase;
        var updated = await Create();
        Check(Format(updated, "select 1;").Contains("SELECT"), "Refresh settings for every invocation.");
        Check(context.Generator.Options.KeywordCasing == KeywordCasing.Lowercase, "Do not mutate earlier context.");
        const string commented = "-- heading\nSELECT a, /* field */ b FROM dbo.t; -- trailing\n";
        foreach (bool native in new[] { false, true })
        foreach (bool axial in new[] { false, true })
        {
            FormatSettingsLoader.Comments = native;
            var commentsContext = await Create();
            var text = Format(commentsContext, commented, new SettingsManager.TSqlCodeFormatSettings { preserveComments = axial });
            Check(Comments(text) == (native || axial ? 3 : 0), "Preserve each comment exactly once.");
            Valid(text, commentsContext);
        }
        FormatSettingsLoader.Comments = true;
        var special = await Create();
        var specialSql = Format(special, "SELECT DISTINCT TOP 5 a, b FROM dbo.t;", new SettingsManager.TSqlCodeFormatSettings
        {
            breakSelectFieldsAfterTopAndUnindent = true
        });
        Check(specialSql.Replace("\r", "").Contains("\n  a"), "Axial SELECT indentation runs after SSMS with SSMS indent size.");
        Valid(specialSql, special);
        const string sample = "-- heading\nCREATE OR ALTER PROCEDURE dbo.p @x int, @y int AS BEGIN DECLARE @a int=1, @b int=2; SELECT DISTINCT TOP 5 t.a, GETDATE(), CASE WHEN t.a=1 THEN 'one' ELSE 'other' END AS label FROM dbo.t AS t INNER JOIN dbo.u AS u ON t.a=u.a CROSS JOIN dbo.v AS v; SELECT 2; END;";
        foreach (var flag in typeof(SettingsManager.TSqlCodeFormatSettings).GetFields())
        {
            var settings = new SettingsManager.TSqlCodeFormatSettings();
            flag.SetValue(settings, true);
            var result = Format(await Create(), sample, settings);
            Valid(result, special);
            Check(Comments(result) == 1, "Axial option retains native comments: " + flag.Name);
        }
        var allOptions = new SettingsManager.TSqlCodeFormatSettings();
        foreach (var flag in typeof(SettingsManager.TSqlCodeFormatSettings).GetFields()) flag.SetValue(allOptions, true);
        var combined = Format(await Create(), sample, allOptions);
        Valid(combined, special);
        Check(Comments(combined) == 1, "Combined Axial options retain comments.");
        await Reject(() => Task.FromResult(Format(special, "SELECT FROM ;")), "syntax error");
        SqlFormatterExtension.ExtensibilityInstance = null;
        await Reject(() => Create(), "has not initialized");
        SqlFormatterExtension.ExtensibilityInstance = new object();
        FormatSettingsLoader.Source = ConfigSource.Defaults;
        await Reject(() => Create(), "could not load");
        FormatSettingsLoader.Source = ConfigSource.Mixed;
        await Create(buffer);
        FormatSettingsLoader.Source = ConfigSource.UnifiedSettings;
        FormatSettingsLoader.FailSynchronously = true;
        await Reject(() => Create(), "settings failed synchronously");
        FormatSettingsLoader.FailSynchronously = false;
        FormatSettingsLoader.FailAsynchronously = true;
        await Reject(() => Create(), "settings failed asynchronously");
        FormatSettingsLoader.FailAsynchronously = false;
        await Reject(() => Create(token: new CancellationToken(true)), "canceled");
        await Reject(() => SsmsFormatterReflection.CreateAsync(typeof(string).Assembly, null, default), "Missing type");
        Console.WriteLine($"Passed {assertions} SSMS formatter integration assertions.");
    }
}
