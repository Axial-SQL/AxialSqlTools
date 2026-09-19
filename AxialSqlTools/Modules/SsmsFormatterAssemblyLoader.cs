using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace AxialSqlTools
{
    internal static class SsmsFormatterAssemblyLoader
    {
        internal static Assembly Load(string ideDirectory)
        {
            // Reuse SSMS's assembly and load context if the native extension is already active.
            var loaded = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(assembly =>
                string.Equals(assembly.GetName().Name, SsmsFormatterReflection.AssemblyName, StringComparison.OrdinalIgnoreCase));
            if (loaded != null)
                return loaded;

            if (string.IsNullOrEmpty(ideDirectory))
                throw new InvalidOperationException("Unable to locate the running SSMS installation.");

            var path = Path.Combine(ideDirectory, "Extensions", "Microsoft", "SSMS.SqlFormatter",
                SsmsFormatterReflection.AssemblyName + ".dll");
            if (!File.Exists(path))
                throw new FileNotFoundException("The SQL Formatter DLL was not found in the running SSMS installation. " +
                    "Install an SSMS version that includes SQL Formatter.", path);

            // LoadFrom uses the installed assembly identity and permits dependencies alongside
            // it to resolve. Do not distribute SSMS DLLs or pin the build machine's version.
            return Assembly.LoadFrom(path);
        }
    }
}
