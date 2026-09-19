using Microsoft.VisualStudio.Shell;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace AxialSqlTools
{
    internal static class SsmsFormatterHost
    {
        internal static SsmsFormatterContext Create()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // Query text stays on the existing DTE path. Load SSMS global settings without
            // resolving a managed editor buffer or loading additional editor assemblies.
            return ThreadHelper.JoinableTaskFactory.Run(() => CreateAsync(null, CancellationToken.None));
        }

        internal static async Task<SsmsFormatterContext> CreateAsync(object textBuffer, CancellationToken cancellationToken)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            try
            {
                var assembly = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(
                    a => a.GetName().Name == SsmsFormatterReflection.AssemblyName);
                if (assembly == null)
                {
                    // Load only from the running SSMS installation, never from a query/project folder.
                    var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Extensions", "Microsoft",
                        "SSMS.SqlFormatter", SsmsFormatterReflection.AssemblyName + ".dll");
                    if (!File.Exists(path))
                        throw new FileNotFoundException("The SSMS SQL Formatter is not installed in this SSMS version.");
                    assembly = Assembly.LoadFrom(path);
                }
                return await SsmsFormatterReflection.CreateAsync(assembly, textBuffer, cancellationToken);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                // Exception messages from the host can include query text. Log only the failure type.
                AxialSqlToolsPackage._logger.Warn("Unable to initialize the SSMS formatter ({0}).", ex.GetType().Name);
                throw new InvalidOperationException("Unable to use the SSMS SQL Formatter. " + ex.Message, ex);
            }
        }

    }
}
