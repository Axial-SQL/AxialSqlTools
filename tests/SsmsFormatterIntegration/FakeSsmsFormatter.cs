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
        internal static SqlScriptGeneratorOptions ToScriptGeneratorOptions(FormatSettings settings) => settings.Options;
        private static TSqlParser CreateParser(SqlVersion version, SqlEngineType engine) => new TSql160Parser(true, engine);
        private static SqlScriptGenerator CreateScriptGenerator(SqlScriptGeneratorOptions options) => new Sql160ScriptGenerator(options);
    }
}
