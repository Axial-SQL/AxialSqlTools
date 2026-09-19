using Microsoft.SqlServer.TransactSql.ScriptDom;
using System;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace AxialSqlTools
{
    internal sealed class SsmsFormatterContext
    {
        internal TSqlParser Parser { get; }
        internal SqlScriptGenerator Generator { get; }

        internal SsmsFormatterContext(TSqlParser parser, SqlScriptGenerator generator)
        {
            Parser = parser;
            Generator = generator;
        }
    }

    // SSMS owns these internal APIs. Keep their reflection contract in one place.
    // Never reuse FormatSettings.Instance: it starts with CLR defaults and can be stale.
    internal static class SsmsFormatterReflection
    {
        internal const string AssemblyName = "Microsoft.SqlServer.Management.SqlFormatter";
        private const BindingFlags StaticMembers = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static;

        internal static async Task<SsmsFormatterContext> CreateAsync(
            Assembly assembly, object textBuffer, CancellationToken cancellationToken, bool disregardSsmsSettings = false,
            Func<Type, CancellationToken, Task<object>> resolveExtensibilityService = null)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var extension = RequiredType(assembly, "SqlFormatterExtension");
            var extensibilityProperty = extension.GetProperty("ExtensibilityInstance", StaticMembers)
                ?? throw Incompatible("Missing ExtensibilityInstance property.");
            var extensibility = extensibilityProperty.GetValue(null);
            if (extensibility == null && resolveExtensibilityService != null)
            {
                // Settings can open before SSMS's SQL-content-type extension is activated.
                // Ask the shell for its service using the installed SDK's exact runtime type.
                extensibility = await resolveExtensibilityService(extensibilityProperty.PropertyType, cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                // Prefer the native instance if SSMS initialized while service resolution yielded.
                extensibility = extensibilityProperty.GetValue(null) ?? extensibility;
            }
            if (extensibility == null)
                throw new InvalidOperationException("The SSMS formatter settings service is unavailable. Try again after SSMS finishes starting.");
            if (!extensibilityProperty.PropertyType.IsInstanceOfType(extensibility))
                throw Incompatible("The formatter settings service has an incompatible type.");

            var loader = RequiredType(assembly, "FormatSettingsLoader");
            var load = RequiredMethod(loader, "LoadAsync", 5);
            var sourceType = load.GetParameters()[3].ParameterType;
            var source = Enum.Parse(sourceType, "MenuOrPalette");
            // A buffer allows SSMS to resolve the document's .editorconfig itself.
            var task = Invoke(load, extensibility, null, textBuffer, source, cancellationToken) as Task;
            if (task == null) throw Incompatible("LoadAsync did not return a Task.");
            await task;
            cancellationToken.ThrowIfCancellationRequested();
            var settings = task.GetType().GetProperty("Result")?.GetValue(task);
            if (settings == null) throw Incompatible("LoadAsync did not return settings.");
            var configSource = settings.GetType().GetProperty("ConfigSource")?.GetValue(settings)?.ToString();
            if (configSource != "UnifiedSettings" && configSource != "Mixed")
                throw new InvalidOperationException("SSMS could not load its SQL Formatter settings. Check the SSMS SQL Formatter options and try again.");

            var helper = RequiredType(assembly, "SqlFormatHelper");
            // Always let SSMS apply its settings and choose the actual implementation types first.
            var sharedOptions = Invoke(RequiredMethod(helper, "ToScriptGeneratorOptions", 1), settings) as SqlScriptGeneratorOptions;
            if (sharedOptions == null)
                throw Incompatible("SSMS and Axial SQL Tools are using incompatible ScriptDOM assemblies.");
            var parser = Invoke(RequiredMethod(helper, "CreateParser", 2),
                sharedOptions.SqlVersion, sharedOptions.SqlEngineType) as TSqlParser;
            var generator = Invoke(RequiredMethod(helper, "CreateScriptGenerator", 1), sharedOptions) as SqlScriptGenerator;
            if (parser == null || generator == null)
                throw Incompatible("The parser or generator has an incompatible ScriptDOM type.");
            if (disregardSsmsSettings)
            {
                // Equivalent to new TSql170Parser(false) / new Sql170ScriptGenerator(),
                // using whichever concrete versions SSMS actually returned. Do not copy options.
                parser = (TSqlParser)Activator.CreateInstance(parser.GetType(), new object[] { false });
                generator = (SqlScriptGenerator)Activator.CreateInstance(generator.GetType());
            }
            return new SsmsFormatterContext(parser, generator);
        }

        private static Type RequiredType(Assembly assembly, string name)
        {
            return assembly.GetType(AssemblyName + "." + name, false)
                ?? throw Incompatible("Missing type " + name + ".");
        }

        private static MethodInfo RequiredMethod(Type type, string name, int parameterCount)
        {
            var method = type.GetMethod(name, StaticMembers);
            if (method == null || method.GetParameters().Length != parameterCount)
                throw Incompatible("Missing or changed method " + name + ".");
            return method;
        }

        private static object Invoke(MethodInfo method, params object[] arguments)
        {
            try { return method.Invoke(null, arguments); }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
                throw;
            }
        }

        private static InvalidOperationException Incompatible(string detail)
        {
            return new InvalidOperationException("This SSMS SQL Formatter version is not compatible with Axial SQL Tools. " + detail);
        }
    }
}
