// Handwritten test double for the internal SSMS reflection contract, not SSMS implementation code.
using Microsoft.SqlServer.TransactSql.ScriptDom;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace Microsoft.SqlServer.Management.SqlFormatter
{
    internal static class SqlFormatterExtension
    {
        internal static object ExtensibilityInstance { get; set; } = new object();
    }

    internal enum InvocationSource { MenuOrPalette }
    internal enum ConfigSource { Defaults, UnifiedSettings, EditorConfig, Mixed }

    internal sealed class FormatSettings
    {
        public ConfigSource ConfigSource { get; set; }
        public SqlVersion SqlVersion => Options.SqlVersion;
        internal SqlScriptGeneratorOptions Options { get; set; }
    }

    internal static class FormatSettingsLoader
    {
        internal static KeywordCasing Casing = KeywordCasing.Lowercase;
        internal static bool Comments = true;
        internal static ConfigSource Source = ConfigSource.UnifiedSettings;
        internal static bool FailSynchronously;
        internal static bool FailAsynchronously;
        internal static object LastBuffer;
        internal static int Loads;

        public static Task<FormatSettings> LoadAsync(object extensibility, object textView, object textBuffer,
            InvocationSource invocationSource, CancellationToken cancellationToken)
        {
            if (FailSynchronously) throw new InvalidOperationException("settings failed synchronously");
            return LoadCoreAsync(textBuffer, cancellationToken);
        }

        private static async Task<FormatSettings> LoadCoreAsync(object buffer, CancellationToken token)
        {
            await Task.Yield();
            token.ThrowIfCancellationRequested();
            if (FailAsynchronously) throw new InvalidOperationException("settings failed asynchronously");
            LastBuffer = buffer;
            Loads++;
            return new FormatSettings
            {
                ConfigSource = Source,
                Options = new SqlScriptGeneratorOptions
                {
                    SqlVersion = SqlVersion.Sql160,
                    SqlEngineType = SqlEngineType.All,
                    KeywordCasing = Casing,
                    IndentationSize = buffer == null ? 2 : 6,
                    AlignClauseBodies = true,
                    PreserveComments = Comments
                }
            };
        }
    }

    internal static class SqlFormatHelper
    {
        internal static int OptionsMappings;
        internal static TSqlParser LastParser;
        internal static SqlScriptGenerator LastGenerator;
        // Simulate a factory choosing a newer implementation than the settings enum suggests.
        internal static bool Use170Implementation;
        internal static SqlScriptGeneratorOptions ToScriptGeneratorOptions(FormatSettings settings)
        {
            OptionsMappings++;
            return settings.Options;
        }
        private static TSqlParser CreateParser(SqlVersion version, SqlEngineType engine)
        {
            LastParser = Use170Implementation ? (TSqlParser)new TSql170Parser(true, engine) : new TSql160Parser(true, engine);
            return LastParser;
        }
        private static SqlScriptGenerator CreateScriptGenerator(SqlScriptGeneratorOptions options)
        {
            LastGenerator = Use170Implementation ? (SqlScriptGenerator)new Sql170ScriptGenerator(options) : new Sql160ScriptGenerator(options);
            return LastGenerator;
        }
    }
}
