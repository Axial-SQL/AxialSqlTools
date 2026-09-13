namespace AxialSqlTools
{
    public sealed class ScriptObjectSelectionItem
    {
        public ScriptObjectSelectionItem(string typeDesc, string schemaName, string objectName, int objectId,
            string databaseName, int parentObjectId, string parentObjectName,
            string parentTypeDesc = null, string columnName = null)
        {
            TypeDesc = typeDesc;
            SchemaName = schemaName;
            ObjectName = objectName;
            ObjectId = objectId;
            DatabaseName = databaseName;
            ParentObjectId = parentObjectId;
            ParentObjectName = parentObjectName;
            ParentTypeDesc = parentTypeDesc;
            ColumnName = columnName;
        }

        public string TypeDesc { get; }
        public string SchemaName { get; }
        public string ObjectName { get; }
        public int ObjectId { get; }
        public string DatabaseName { get; }
        public int ParentObjectId { get; }
        public string ParentObjectName { get; }
        public string ParentTypeDesc { get; }
        public string ColumnName { get; }

        public string DisplayName => string.IsNullOrEmpty(ParentObjectName)
            ? $"{DatabaseName}.{SchemaName}.{ObjectName} ({TypeDesc})"
            : $"{DatabaseName}.{SchemaName}.{ParentObjectName}.{ObjectName} ({TypeDesc})";
    }
}
