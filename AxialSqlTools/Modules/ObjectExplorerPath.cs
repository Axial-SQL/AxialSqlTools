using System;
using System.Collections.Generic;

namespace AxialSqlTools
{
    internal static class ObjectExplorerPath
    {
        // Collection and object URNs let SSMS resolve its own localized container folders.
        // Never search by displayed captions such as 'Tables' or 'Programmability'.
        public static IList<string> GetRelativeSteps(ScriptObjectSelectionItem item)
        {
            var steps = new List<string>();
            string path = string.Empty;
            Append(steps, ref path, "Database", item.DatabaseName);
            string type = item.TypeDesc;
            if (IsChild(type))
            {
                if (string.IsNullOrEmpty(item.ParentObjectName))
                    throw new InvalidOperationException("The parent object could not be resolved for " + item.DisplayName + ".");
                Append(steps, ref path, ObjectType(item.ParentTypeDesc ?? "USER_TABLE"), item.ParentObjectName, item.SchemaName);
                if (type == "DEFAULT_CONSTRAINT")
                {
                    if (string.IsNullOrEmpty(item.ColumnName))
                        throw new InvalidOperationException("The column for this default constraint could not be resolved.");
                    Append(steps, ref path, "Column", item.ColumnName);
                    // A column default has no independently selectable OE node.
                    return steps;
                }
                Append(steps, ref path, ObjectType(type), item.ObjectName);
            }
            else
                Append(steps, ref path, ObjectType(type), item.ObjectName, item.SchemaName);
            return steps;
        }

        public static string ServerUrn(string server) => "Server[@Name='" + Escape(server) + "']";

        public static string NormalizeServer(string server)
        {
            // Keep instance names, full DNS names, protocol prefixes, and ports intact.
            // Removing any of those could select a different server's tree.
            if (server == null) return string.Empty;
            var parts = server.Trim().Split(',');
            for (int i = 0; i < parts.Length; i++) parts[i] = parts[i].Trim();
            return string.Join(",", parts);
        }

        public static bool SameConnection(string requestedServer, string requestedUser, bool requestedIntegrated,
            string actualServer, string actualUser, bool actualIntegrated)
        {
            return string.Equals(NormalizeServer(requestedServer), NormalizeServer(actualServer), StringComparison.OrdinalIgnoreCase)
                && requestedIntegrated == actualIntegrated
                && (requestedIntegrated || string.Equals(requestedUser ?? string.Empty, actualUser ?? string.Empty, StringComparison.Ordinal));
        }

        private static void Append(List<string> steps, ref string path, string type, string name, string schema = null)
        {
            path += "/" + type;
            steps.Add(path);
            path += "[@Name='" + Escape(name) + "'";
            if (!string.IsNullOrEmpty(schema)) path += " and @Schema='" + Escape(schema) + "'";
            path += "]";
            steps.Add(path);
        }

        private static string Escape(string value) => value.Replace("'", "''");

        private static bool IsChild(string type) => type == "INDEX" || type == "SQL_TRIGGER"
            || type == "CLR_TRIGGER" || type == "PRIMARY_KEY_CONSTRAINT" || type == "UNIQUE_CONSTRAINT"
            || type == "FOREIGN_KEY_CONSTRAINT" || type == "CHECK_CONSTRAINT" || type == "DEFAULT_CONSTRAINT";

        private static string ObjectType(string type)
        {
            switch (type)
            {
                case "USER_TABLE": case "SYSTEM_TABLE": case "EXTERNAL_TABLE": return "Table";
                case "VIEW": return "View";
                case "SQL_STORED_PROCEDURE": case "CLR_STORED_PROCEDURE": return "StoredProcedure";
                case "SQL_SCALAR_FUNCTION": case "SQL_INLINE_TABLE_VALUED_FUNCTION":
                case "SQL_TABLE_VALUED_FUNCTION": case "CLR_SCALAR_FUNCTION":
                case "CLR_TABLE_VALUED_FUNCTION": return "UserDefinedFunction";
                case "SYNONYM": return "Synonym";
                case "TYPE_TABLE": return "UserDefinedTableType";
                case "SQL_TRIGGER": case "CLR_TRIGGER": return "Trigger";
                case "INDEX": case "PRIMARY_KEY_CONSTRAINT": case "UNIQUE_CONSTRAINT": return "Index";
                case "FOREIGN_KEY_CONSTRAINT": return "ForeignKey";
                case "CHECK_CONSTRAINT": return "Check";
                case "DEFAULT_CONSTRAINT": return "Default";
                case "CLR_AGGREGATE_FUNCTION": return "UserDefinedAggregate";
                case "SEQUENCE_OBJECT": return "Sequence";
                default: throw new NotSupportedException("Open in Object Explorer does not support this object type: " + type + ".");
            }
        }
    }
}
