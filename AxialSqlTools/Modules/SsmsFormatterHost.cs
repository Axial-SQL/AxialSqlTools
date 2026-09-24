using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.Linq;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace AxialSqlTools
{
    internal static class SsmsFormatterHost
    {
        private const string EditorAdapterContract = "Microsoft.VisualStudio.Editor.IVsEditorAdaptersFactoryService";

        internal static SsmsFormatterContext Create(bool disregardSsmsSettings = false)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var buffer = GetActiveBuffer();
            var context = ThreadHelper.JoinableTaskFactory.Run(() => CreateAsync(buffer, CancellationToken.None, disregardSsmsSettings));
            if (!ReferenceEquals(buffer, GetActiveBuffer()))
                throw new InvalidOperationException("The active query changed while loading formatter settings. Try Format again.");
            return context;
        }

        internal static async Task<SsmsFormatterContext> CreateAsync(object textBuffer, CancellationToken cancellationToken, bool disregardSsmsSettings = false)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            try
            {
                // The formatter lives outside the normal CLR probing path. Resolve it from
                // this SSMS process, not the version/path used to build the extension.
                var assembly = LoadFormatterAssembly();
                return await SsmsFormatterReflection.CreateAsync(assembly, textBuffer, cancellationToken, disregardSsmsSettings,
                    ResolveExtensibilityServiceAsync);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                // Exception messages from the host can include query text. Log only the failure type.
                AxialSqlToolsPackage._logger.Warn("Unable to initialize the SSMS formatter ({0}).", ex.GetType().Name);
                throw new InvalidOperationException("Unable to use the SSMS SQL Formatter. " + ex.Message, ex);
            }
        }

        internal static async Task<SsmsFormatterSettingsSnapshot> ReadSettingsAsync(CancellationToken cancellationToken)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            var assembly = LoadFormatterAssembly();
            // No document buffer: report SSMS's global options, independently of the active
            // editor, .editorconfig overrides, and Axial's Disregard checkbox.
            var settings = await SsmsFormatterReflection.LoadSettingsAsync(assembly, null,
                cancellationToken, ResolveExtensibilityServiceAsync);
            return SsmsFormatterSettingsSnapshot.Capture(settings, assembly);
        }

        private static System.Reflection.Assembly LoadFormatterAssembly()
        {
            using (var process = Process.GetCurrentProcess())
                return SsmsFormatterAssemblyLoader.Load(Path.GetDirectoryName(process.MainModule.FileName));
        }

        private static async Task<object> ResolveExtensibilityServiceAsync(Type serviceType, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            // VisualStudioExtensibility is a registered shell service available to VSSDK packages,
            // independently of the SQL formatter extension's lazy activation. Do not set its statics.
            var service = await AsyncServiceProvider.GlobalProvider.GetServiceAsync(serviceType);
            cancellationToken.ThrowIfCancellationRequested();
            return service;
        }

        private static object GetActiveBuffer()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var manager = Package.GetGlobalService(typeof(SVsTextManager)) as IVsTextManager;
            if (manager == null || manager.GetActiveView(0, null, out var view) < 0 || view == null
                || view.GetBuffer(out var nativeBuffer) < 0 || nativeBuffer == null)
                throw new InvalidOperationException("No active SQL editor buffer is available for formatting.");

            var componentModel = Package.GetGlobalService(typeof(SComponentModel)) as IComponentModel;
            if (componentModel == null)
                throw new InvalidOperationException("The SSMS editor component service is unavailable.");

            // Resolve SSMS's existing export by contract. Its runtime interface supplies the method,
            // so there is no Assembly.Load or guessed identity for Microsoft.VisualStudio.Editor.
            var adapter = componentModel.DefaultExportProvider.GetExportedValue<object>(EditorAdapterContract);
            var adapterType = adapter.GetType().GetInterfaces().SingleOrDefault(t => t.FullName == EditorAdapterContract);
            var getBuffer = adapterType?.GetMethod("GetDocumentBuffer")
                ?? throw new InvalidOperationException("This SSMS version does not expose the expected editor buffer API.");
            return getBuffer.Invoke(adapter, new object[] { nativeBuffer })
                ?? throw new InvalidOperationException("The active query has no managed editor buffer.");
        }
    }
}
