using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TextManager.Interop;
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
        internal static SsmsFormatterContext CreateForActiveDocument()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var buffer = GetActiveBuffer();
            var context = ThreadHelper.JoinableTaskFactory.Run(() => CreateAsync(buffer, CancellationToken.None));
            // Settings loading yields to the shell. Do not apply one document's settings to another.
            if (!ReferenceEquals(buffer, GetActiveBuffer()))
                throw new InvalidOperationException("The active query changed while loading formatter settings. Try Format again.");
            return context;
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

        private static object GetActiveBuffer()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var manager = Package.GetGlobalService(typeof(SVsTextManager)) as IVsTextManager;
            if (manager == null || manager.GetActiveView(0, null, out var view) < 0 || view == null
                || view.GetBuffer(out var nativeBuffer) < 0 || nativeBuffer == null)
                throw new InvalidOperationException("No active SQL editor buffer is available for formatting.");

            // Use the shell's MEF service without adding/shipping another editor SDK dependency.
            var componentAssembly = Assembly.Load("Microsoft.VisualStudio.ComponentModelHost");
            var serviceType = componentAssembly.GetType("Microsoft.VisualStudio.ComponentModelHost.SComponentModel", true);
            var componentType = componentAssembly.GetType("Microsoft.VisualStudio.ComponentModelHost.IComponentModel", true);
            var componentModel = Package.GetGlobalService(serviceType);
            if (componentModel == null)
                throw new InvalidOperationException("The SSMS editor component service is unavailable.");
            var adapterType = Assembly.Load("Microsoft.VisualStudio.Editor")
                .GetType("Microsoft.VisualStudio.Editor.IVsEditorAdaptersFactoryService", true);
            var getService = componentType.GetMethods().Single(m => m.Name == "GetService"
                && m.IsGenericMethodDefinition && m.GetParameters().Length == 0);
            var adapter = getService.MakeGenericMethod(adapterType).Invoke(componentModel, null);
            if (adapter == null)
                throw new InvalidOperationException("The SSMS editor adapter service is unavailable.");
            var getBuffer = adapterType.GetMethod("GetDocumentBuffer")
                ?? throw new InvalidOperationException("This SSMS version does not expose the expected editor buffer API.");
            var buffer = getBuffer.Invoke(adapter, new object[] { nativeBuffer });
            return buffer ?? throw new InvalidOperationException("The active query has no managed editor buffer.");
        }
    }
}
