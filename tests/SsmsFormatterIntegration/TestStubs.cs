// Minimal dependencies for running the production formatting pipeline outside SSMS.
using System;
namespace AxialSqlTools
{
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
                // preserveComments option is not included because it belongs to a separate code branch.
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
    public static class AxialSqlToolsPackage
    {
        public static readonly TestLogger _logger = new TestLogger();
        public sealed class TestLogger
        {
            public void Warn(string message, params object[] args) { }
            public void Error(Exception ex, string message) { throw new Exception(message, ex); }
        }
    }
}
