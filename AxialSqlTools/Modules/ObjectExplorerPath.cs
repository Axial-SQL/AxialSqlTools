using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace AxialSqlTools
{
    internal sealed class ObjectExplorerTreeStep
    {
        public ObjectExplorerTreeStep(string type, string name, string schema, params string[] folders)
        {
            Type = type;
            Name = name;
            Schema = schema;
            Folders = folders;
        }

        public string Type { get; }
        public string Name { get; }
        public string Schema { get; }
        public string[] Folders { get; }
        public string DisplayName => string.IsNullOrEmpty(Schema) ? Name : Schema + "." + Name;

        public bool Matches(IObjectExplorerTreeNode node, bool withinSchema)
        {
            if (node.IsFolder) return false;
            if (!string.IsNullOrEmpty(node.UrnPath) && !node.UrnPath.EndsWith("/" + Type, StringComparison.Ordinal)
                && node.UrnPath != Type) return false;
            // Use catalog identity when SSMS supplies it. Flattened captions can collide:
            // [a.b].[c] and [a].[b.c] both display as a.b.c.
            if (!string.IsNullOrEmpty(node.NavigationContext))
            {
                var identity = Regex.Match(node.NavigationContext, @"(?:^|/)" + Regex.Escape(Type)
                    + @"\[(?<attributes>(?:[^\]']|'(?:''|[^'])*')*)\]$");
                if (identity.Success)
                {
                    string actualName = null, actualSchema = null;
                    foreach (Match attribute in Regex.Matches(identity.Groups["attributes"].Value,
                        @"@(?<key>Name|Schema)='(?<value>(?:''|[^'])*)'"))
                    {
                        string value = attribute.Groups["value"].Value.Replace("''", "'");
                        if (attribute.Groups["key"].Value == "Name") actualName = value;
                        else actualSchema = value;
                    }
                    if (actualName != null)
                        return actualName == Name && (string.IsNullOrEmpty(Schema) || actualSchema == Schema);
                }
            }
            // Resolver results already have the catalog's exact casing. Do not collapse two
            // distinct objects in a case-sensitive database, or split identifiers on dots.
            if (!string.IsNullOrEmpty(Schema))
            {
                if (node.InvariantName == DisplayName || node.Name == DisplayName) return true;
                return withinSchema && (node.InvariantName == Name || node.Name == Name);
            }
            return node.Name == Name || node.InvariantName == Name;
        }

        public bool IsSchemaFolder(IObjectExplorerTreeNode node) => (node.IsFolder
            || (node.UrnPath?.EndsWith("/Schema", StringComparison.Ordinal) ?? false))
            && !string.IsNullOrEmpty(Schema) && (node.Name == Schema || node.InvariantName == Schema);

        public bool IsContainer(IObjectExplorerTreeNode node)
        {
            if (!node.IsFolder) return false;
            if (node.UrnPath?.EndsWith("/" + Type, StringComparison.Ordinal) == true) return true;
            if (!string.IsNullOrEmpty(Schema) && FolderMatches(node, "Schemas")) return true;
            return Folders.Any(folder => FolderMatches(node, folder));
        }

        private static bool FolderMatches(IObjectExplorerTreeNode node, string name)
        {
            // SQL Search uses containedItem.context["UniqueName"], not localized captions.
            string unique = node.UniqueName;
            if (name == "UserProgrammability" && unique != null
                && unique.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (string.Equals(unique, name, StringComparison.OrdinalIgnoreCase)) return true;
            // Older/synthetic folders may have no context. Captions are a last resort only.
            return string.Equals(node.InvariantName, name, StringComparison.OrdinalIgnoreCase)
                || string.Equals(node.Name, name, StringComparison.OrdinalIgnoreCase);
        }
    }

    internal static class ObjectExplorerPath
    {
        public static IList<ObjectExplorerTreeStep> GetTreeSteps(ScriptObjectSelectionItem item)
        {
            if (item == null) throw new ArgumentNullException(nameof(item));
            var steps = new List<ObjectExplorerTreeStep>
            {
                new ObjectExplorerTreeStep("Database", item.DatabaseName, null,
                    "Databases", "SystemDatabases", "System Databases")
            };
            string type = item.TypeDesc;
            if (IsChild(type))
            {
                if (string.IsNullOrEmpty(item.ParentObjectName))
                    throw new InvalidOperationException("The parent object could not be resolved for " + item.DisplayName + ".");
                steps.Add(ObjectStep(item.ParentTypeDesc ?? "USER_TABLE", item.ParentObjectName, item.SchemaName));
                switch (type)
                {
                    case "DEFAULT_CONSTRAINT":
                        if (string.IsNullOrEmpty(item.ColumnName))
                            throw new InvalidOperationException("The column for this default constraint could not be resolved.");
                        steps.Add(new ObjectExplorerTreeStep("Column", item.ColumnName, null, "Columns"));
                        break;
                    case "PRIMARY_KEY_CONSTRAINT":
                        steps.Add(new ObjectExplorerTreeStep("Index", item.ObjectName, null, "Keys", "Indexes"));
                        break;
                    case "FOREIGN_KEY_CONSTRAINT":
                        steps.Add(new ObjectExplorerTreeStep("ForeignKey", item.ObjectName, null, "Keys"));
                        break;
                    case "CHECK_CONSTRAINT":
                        steps.Add(new ObjectExplorerTreeStep("Check", item.ObjectName, null, "Constraints", "Checks"));
                        break;
                    case "SQL_TRIGGER": case "CLR_TRIGGER":
                        steps.Add(new ObjectExplorerTreeStep("Trigger", item.ObjectName, null, "Triggers"));
                        break;
                    default:
                        steps.Add(new ObjectExplorerTreeStep("Index", item.ObjectName, null, "Indexes"));
                        break;
                }
            }
            else steps.Add(ObjectStep(type, item.ObjectName, item.SchemaName));
            return steps;
        }

        private static ObjectExplorerTreeStep ObjectStep(string type, string name, string schema)
        {
            switch (type)
            {
                case "USER_TABLE": case "SYSTEM_TABLE": case "EXTERNAL_TABLE":
                    return new ObjectExplorerTreeStep("Table", name, schema, "UserTables", "Tables",
                        "SystemTables", "System Tables", "ExternalTables", "External Tables");
                case "VIEW":
                    return new ObjectExplorerTreeStep("View", name, schema, "Views", "SystemViews", "System Views");
                case "SQL_STORED_PROCEDURE": case "CLR_STORED_PROCEDURE":
                    return new ObjectExplorerTreeStep("StoredProcedure", name, schema, "UserProgrammability", "Programmability",
                        "StoredProcedures", "Stored Procedures", "SystemStoredProcedures", "System Stored Procedures");
                case "SQL_SCALAR_FUNCTION": case "CLR_SCALAR_FUNCTION":
                    return Function(name, schema, "Scalar-valuedFunctions", "Scalar-valued Functions");
                case "SQL_INLINE_TABLE_VALUED_FUNCTION": case "SQL_TABLE_VALUED_FUNCTION": case "CLR_TABLE_VALUED_FUNCTION":
                    return Function(name, schema, "Table-valuedFunctions", "Table-valued Functions");
                case "SYNONYM": return new ObjectExplorerTreeStep("Synonym", name, schema, "Synonyms");
                case "TYPE_TABLE":
                    return new ObjectExplorerTreeStep("UserDefinedTableType", name, schema, "UserProgrammability", "Programmability",
                        "Types", "UserDefinedTableTypes", "User-Defined Table Types");
                case "CLR_AGGREGATE_FUNCTION":
                    return new ObjectExplorerTreeStep("UserDefinedAggregate", name, schema, "UserProgrammability", "Programmability",
                        "UsrDbFunctions", "Functions", "AggregateFunctions", "Aggregate Functions");
                case "SEQUENCE_OBJECT":
                    return new ObjectExplorerTreeStep("Sequence", name, schema, "UserProgrammability", "Programmability", "Sequences");
                default: throw new NotSupportedException("Open in Object Explorer does not support this object type: " + type + ".");
            }
        }

        private static ObjectExplorerTreeStep Function(string name, string schema, string folder, string caption) =>
            new ObjectExplorerTreeStep("UserDefinedFunction", name, schema, "UserProgrammability", "Programmability",
                "UsrDbFunctions", "Functions", folder, caption, "SystemFunctions", "System Functions");

        private static bool IsChild(string type) => type == "INDEX" || type == "SQL_TRIGGER"
            || type == "CLR_TRIGGER" || type == "PRIMARY_KEY_CONSTRAINT" || type == "UNIQUE_CONSTRAINT"
            || type == "FOREIGN_KEY_CONSTRAINT" || type == "CHECK_CONSTRAINT" || type == "DEFAULT_CONSTRAINT";

        public static bool SameAuthentication(string requested, string actual)
        {
            // SqlClient and SMO both use NotSpecified for legacy SQL password connections.
            if (requested == "NotSpecified") requested = "SqlPassword";
            if (actual == "NotSpecified") actual = "SqlPassword";
            return string.Equals(requested, actual, StringComparison.Ordinal);
        }

        public static string NormalizeServer(string server)
        {
            if (server == null) return string.Empty;
            var parts = server.Trim().Split(',');
            for (int i = 0; i < parts.Length; i++) parts[i] = parts[i].Trim();
            return string.Join(",", parts);
        }

        public static bool SameConnection(string requestedServer, string requestedUser, bool requestedIntegrated,
            string actualServer, string actualUser, bool actualIntegrated) =>
            string.Equals(NormalizeServer(requestedServer), NormalizeServer(actualServer), StringComparison.OrdinalIgnoreCase)
            && requestedIntegrated == actualIntegrated
            && (requestedIntegrated || string.Equals(requestedUser ?? string.Empty, actualUser ?? string.Empty, StringComparison.Ordinal));
    }
}
