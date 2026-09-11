using System;
namespace AxialSqlTools
{
    // The parser only needs the package's logger. The data models are linked production code.
    public sealed partial class AxialSqlToolsPackage
    {
        public static TestLogger _logger;
        public sealed class TestLogger { public void Error(Exception error, string message) { } }
    }
}
