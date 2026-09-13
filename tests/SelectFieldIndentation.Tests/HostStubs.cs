using System;
using System.Collections.Generic;

namespace AxialSqlTools
{
    // Only the host logger and unrelated connection storage are replaced.
    // The formatter, comment interleaver and JSON settings model are production code.
    public static class AxialSqlToolsPackage
    {
        public static TestLogger _logger = new TestLogger();
        public class TestLogger
        {
            public void Error(Exception error, string message) { throw new Exception(message, error); }
        }
    }
    internal static class SavedConnectionStore
    {
        public static List<SettingsManager.DataTransferSavedConnection> Load() { throw new NotSupportedException(); }
        public static void Save(List<SettingsManager.DataTransferSavedConnection> connections) { throw new NotSupportedException(); }
    }
}
