using Microsoft.SqlServer.TransactSql.ScriptDom;

namespace AxialSqlTools
{
    public static class TSqlParserFactory
    {
        public static TSqlParser Create(bool initialQuotedIdentifiers)
        {
            return Create(SettingsManager.GetTSqlParserVersion(), initialQuotedIdentifiers);
        }

        public static TSqlParser Create(SettingsManager.TSqlParserVersion version, bool initialQuotedIdentifiers)
        {
            switch (version)
            {
                case SettingsManager.TSqlParserVersion.Sql80:
                    return new TSql80Parser(initialQuotedIdentifiers);
                case SettingsManager.TSqlParserVersion.Sql90:
                    return new TSql90Parser(initialQuotedIdentifiers);
                case SettingsManager.TSqlParserVersion.Sql100:
                    return new TSql100Parser(initialQuotedIdentifiers);
                case SettingsManager.TSqlParserVersion.Sql110:
                    return new TSql110Parser(initialQuotedIdentifiers);
                case SettingsManager.TSqlParserVersion.Sql120:
                    return new TSql120Parser(initialQuotedIdentifiers);
                case SettingsManager.TSqlParserVersion.Sql130:
                    return new TSql130Parser(initialQuotedIdentifiers);
                case SettingsManager.TSqlParserVersion.Sql140:
                    return new TSql140Parser(initialQuotedIdentifiers);
                case SettingsManager.TSqlParserVersion.Sql150:
                    return new TSql150Parser(initialQuotedIdentifiers);
                case SettingsManager.TSqlParserVersion.Sql160:
                    return new TSql160Parser(initialQuotedIdentifiers);
                case SettingsManager.TSqlParserVersion.Sql170:
                default:
                    return new TSql170Parser(initialQuotedIdentifiers);
            }
        }

        public static SqlScriptGenerator CreateScriptGenerator()
        {
            switch (SettingsManager.GetTSqlParserVersion())
            {
                case SettingsManager.TSqlParserVersion.Sql80:
                    return new Sql80ScriptGenerator();
                case SettingsManager.TSqlParserVersion.Sql90:
                    return new Sql90ScriptGenerator();
                case SettingsManager.TSqlParserVersion.Sql100:
                    return new Sql100ScriptGenerator();
                case SettingsManager.TSqlParserVersion.Sql110:
                    return new Sql110ScriptGenerator();
                case SettingsManager.TSqlParserVersion.Sql120:
                    return new Sql120ScriptGenerator();
                case SettingsManager.TSqlParserVersion.Sql130:
                    return new Sql130ScriptGenerator();
                case SettingsManager.TSqlParserVersion.Sql140:
                    return new Sql140ScriptGenerator();
                case SettingsManager.TSqlParserVersion.Sql150:
                    return new Sql150ScriptGenerator();
                case SettingsManager.TSqlParserVersion.Sql160:
                    return new Sql160ScriptGenerator();
                case SettingsManager.TSqlParserVersion.Sql170:
                default:
                    return new Sql170ScriptGenerator();
            }
        }
    }
}
