// The harness links production formatter files directly; only SSMS host services are stubbed.
namespace AxialSqlTools
{
    public static class AxialSqlToolsPackage
    {
        public static readonly TestLogger _logger = new TestLogger();
        public sealed class TestLogger
        {
            public void Error(System.Exception exception, string message) => throw new System.InvalidOperationException(message, exception);
        }
    }
    public static class SettingsManager
    {
        public static TSqlCodeFormatSettings GetTSqlCodeFormatSettings() => new TSqlCodeFormatSettings();
        public class TSqlCodeFormatSettings
        {
            public bool disregardSsmsFormatterSettings = false;
            public bool preserveComments = false;
            public bool removeNewLineAfterJoin = false;
            public bool addTabAfterJoinOn = false;
            public bool moveCrossJoinToNewLine = false;
            public bool formatCaseAsMultiline = false;
            public bool addNewLineBetweenStatementsInBlocks = false;
            public bool breakSprocParametersPerLine = false;
            public bool uppercaseBuiltInFunctions = false;
            public bool unindentBeginEndBlocks = false;
            public bool breakVariableDefinitionsPerLine = false;
            public bool breakSprocDefinitionParametersPerLine = false;
            // Retain the existing serialized name; this option now handles DISTINCT as well as TOP.
            public bool breakSelectFieldsAfterTopAndUnindent = false;

            public bool HasAnyFormattingEnabled()
            {
                // Formatter selection and comment preservation are handled before post-processing.
                return removeNewLineAfterJoin
                    || addTabAfterJoinOn
                    || moveCrossJoinToNewLine
                    || formatCaseAsMultiline
                    || addNewLineBetweenStatementsInBlocks
                    || breakSprocParametersPerLine
                    || uppercaseBuiltInFunctions
                    || unindentBeginEndBlocks
                    || breakVariableDefinitionsPerLine
                    || breakSprocDefinitionParametersPerLine
                    || breakSelectFieldsAfterTopAndUnindent;
            }
        }

    }
}
