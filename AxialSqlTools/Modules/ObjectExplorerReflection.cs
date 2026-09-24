using System;
using System.Reflection;
using System.Runtime.ExceptionServices;

namespace AxialSqlTools
{
    // The tree exposes its hierarchy item and context through private inherited members.
    // Keep reflection in one place so API changes produce a useful error, not a timeout.
    internal static class ObjectExplorerReflection
    {
        private const BindingFlags Members = BindingFlags.Instance | BindingFlags.Public
            | BindingFlags.NonPublic | BindingFlags.DeclaredOnly;

        public static object Read(object instance, string name)
        {
            for (Type type = instance?.GetType(); type != null; type = type.BaseType)
            {
                var property = type.GetProperty(name, Members);
                if (property != null && property.GetIndexParameters().Length == 0)
                    return property.GetValue(instance);
                var field = type.GetField(name, Members);
                if (field != null) return field.GetValue(instance);
            }
            return null;
        }

        public static string UniqueName(object treeNode)
        {
            var item = Read(treeNode, "containedItem");
            var context = Read(item, "context");
            var getter = Method(context, "get_Item", typeof(string));
            return getter == null ? null : Invoke(getter, context, new object[] { "UniqueName" }) as string;
        }

        public static void EnumerateChildren(object treeNode)
        {
            var method = Method(treeNode, "EnumerateChildren", typeof(bool));
            if (method == null)
                throw new NotSupportedException("This SSMS version does not expose Object Explorer's EnumerateChildren(bool) method.");
            // false means synchronous enumeration, as used by SQL Search. Expand alone does
            // not guarantee that a cold folder's child nodes exist yet.
            Invoke(method, treeNode, new object[] { false });
        }

        public static bool TryShow(object explorer)
        {
            var method = Method(explorer, "Show");
            if (method == null) return false;
            Invoke(method, explorer, new object[0]);
            return true;
        }

        private static MethodInfo Method(object instance, string name, params Type[] arguments)
        {
            for (Type type = instance?.GetType(); type != null; type = type.BaseType)
            {
                var method = type.GetMethod(name, Members, null, arguments, null);
                if (method != null) return method;
            }
            return null;
        }

        private static object Invoke(MethodInfo method, object instance, object[] arguments)
        {
            try { return method.Invoke(instance, arguments); }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
                throw;
            }
        }
    }
}
