using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace AxialSqlTools.DataCompare
{
    public static class SqlText
    {
        public static string Quote(string name) => "[" + (name ?? throw new ArgumentNullException(nameof(name))).Replace("]", "]]") + "]";
        public static string String(string value) => "N'" + value.Replace("'", "''") + "'";
        public static bool IsText(string type) => new[] { "char", "varchar", "nchar", "nvarchar", "text", "ntext", "xml" }.Contains(type);
        public static bool IsBinary(string type) => new[] { "binary", "varbinary", "image", "timestamp", "rowversion" }.Contains(type);
        public static string Hex(byte[] bytes) => "0x" + BitConverter.ToString(bytes).Replace("-", "");
        private static byte[] Utf16(string text)
        {
            var bytes = new byte[text.Length * 2];
            Buffer.BlockCopy(text.ToCharArray(), 0, bytes, 0, bytes.Length);
            return bytes;
        }
        public static string Literal(object value)
        {
            if (ValueCodec.IsNull(value)) return "NULL";
            if (value is string text) return "CONVERT(nvarchar(max), " + Hex(Utf16(text)) + ")";
            if (value is byte[] bytes) return Hex(bytes);
            if (value is bool flag) return flag ? "1" : "0";
            if (value is DateTime date) return "'" + date.ToString("yyyy-MM-ddTHH:mm:ss.fffffff", CultureInfo.InvariantCulture) + "'";
            if (value is DateTimeOffset offset) return "'" + offset.ToString("o", CultureInfo.InvariantCulture) + "'";
            if (value is TimeSpan time) return "'" + time.ToString("c", CultureInfo.InvariantCulture) + "'";
            if (value is Guid guid) return "'" + guid.ToString("D") + "'";
            if (value is SqlDecimal number) return number.ToString();
            if (value is double f64)
            {
                if (double.IsNaN(f64) || double.IsInfinity(f64)) throw new InvalidOperationException("Non-finite SQL floating point value.");
                return "CONVERT(float, '" + f64.ToString("R", CultureInfo.InvariantCulture) + "')";
            }
            if (value is float f32)
            {
                if (float.IsNaN(f32) || float.IsInfinity(f32)) throw new InvalidOperationException("Non-finite SQL floating point value.");
                return "CONVERT(real, '" + f32.ToString("R", CultureInfo.InvariantCulture) + "')";
            }
            if (value is decimal || value is long || value is int || value is short || value is byte)
                return Convert.ToString(value, CultureInfo.InvariantCulture);
            throw new NotSupportedException("Cannot script value type " + value.GetType().Name);
        }

        public static string Literal(object value, TableColumn column)
        {
            if (ValueCodec.IsNull(value)) return "NULL";
            if (value is DateTime date)
            {
                string format = column.Type == "datetime" || column.Type == "smalldatetime" ? "yyyy-MM-ddTHH:mm:ss.fff" : "yyyy-MM-ddTHH:mm:ss.fffffff";
                string type = column.Type == "datetime2" ? "datetime2(" + column.Scale + ")" : column.Type;
                return "CONVERT(" + type + ", '" + date.ToString(format, CultureInfo.InvariantCulture) + "', 126)";
            }
            if (column.Type == "xml") return "CONVERT(xml, " + Literal(value) + ", 1)";
            return Literal(value);
        }

        public static string Equal(TableColumn column, object value)
        {
            string name = Quote(column.Name);
            if (ValueCodec.IsNull(value)) return name + " IS NULL";
            if (value is string text)
            {
                string expression = "CONVERT(varbinary(max), CONVERT(nvarchar(max), " + name + "))";
                return "(" + expression + " = " + Hex(Utf16(text)) + " AND DATALENGTH(" + expression + ") = " + text.Length * 2L + ")";
            }
            if (value is byte[] bytes)
                return "(CONVERT(varbinary(max), " + name + ") = " + Hex(bytes) + " AND DATALENGTH(" + name + ") = " + bytes.Length + ")";
            return name + " = " + Literal(value, column);
        }
    }

    public static class SynchronizationScript
    {
        public static void Validate(ComparisonResult result)
        {
            if (result == null || result.SelectedCount == 0) throw new InvalidOperationException("Select at least one changed row to synchronize.");
            result.Plan.Validate();
            if (result.Plan.Target.HistoryTable) throw new InvalidOperationException("A temporal history table cannot be a synchronization target.");
            if (result.Selection.Count(DifferenceKind.OnlySource, result.Count(DifferenceKind.OnlySource)) > 0)
            {
                var mapped = new HashSet<string>(result.Plan.Columns.Where(c => c.Target.Writable).Select(c => c.Target.Name), StringComparer.Ordinal);
                var missing = result.Plan.Target.Columns.Where(c => c.Writable && !c.Identity && !c.Nullable && !c.HasDefault && !mapped.Contains(c.Name)).ToList();
                if (missing.Count > 0) throw new InvalidOperationException("Inserts need values for required target columns: " + string.Join(", ", missing.Select(c => c.Name)) + ". Map them or deselect source-only rows.");
            }
        }

        public static bool NeedsIdentityInsert(ComparisonResult result) =>
            result.Selection.Count(DifferenceKind.OnlySource, result.Count(DifferenceKind.OnlySource)) > 0 &&
            result.Plan.Columns.Any(c => c.Target.Identity);

        public static string DestinationGuard(ComparisonPlan plan)
        {
            return "IF SERVERPROPERTY('ServerName') IS NULL OR DB_NAME() IS NULL\n" +
                "   OR CONVERT(varbinary(max), CONVERT(nvarchar(256), SERVERPROPERTY('ServerName'))) <> CONVERT(varbinary(max), " + SqlText.String(plan.Target.Server) + ")\n" +
                "   OR CONVERT(varbinary(max), DB_NAME()) <> CONVERT(varbinary(max), " + SqlText.String(plan.Target.Database) + ")\n" +
                "    THROW 51000, 'Wrong target server or database. Connect to the target used for this comparison.', 1;\n";
        }

        public static string SchemaGuard(ComparisonPlan plan)
        {
            var script = new StringBuilder();
            string objectName = SqlText.String(plan.Target.QualifiedName);
            foreach (var mapping in plan.Columns)
            {
                var c = mapping.Target;
                script.Append("IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID(").Append(objectName).Append(", N'U') AND ")
                    .Append("CONVERT(varbinary(max), name)=CONVERT(varbinary(max), ").Append(SqlText.String(c.Name)).Append(") AND ")
                    .Append("TYPE_NAME(system_type_id)=").Append(SqlText.String(c.Type)).Append(" AND max_length=").Append(c.Length)
                    .Append(" AND precision=").Append(c.Precision).Append(" AND scale=").Append(c.Scale)
                    .Append(" AND is_nullable=").Append(c.Nullable ? 1 : 0).Append(" AND is_identity=").Append(c.Identity ? 1 : 0)
                    .Append(" AND is_computed=").Append(c.Computed ? 1 : 0).Append(" AND ").Append(c.Generated ? "generated_always_type<>0" : "generated_always_type=0")
                    .Append(" AND ").Append(c.Collation == null ? "collation_name IS NULL" : "CONVERT(varbinary(max), collation_name)=CONVERT(varbinary(max), " + SqlText.String(c.Collation) + ")")
                    .Append(" AND encryption_type IS NULL)\n    THROW 51001, 'Target column metadata changed. Run the comparison again.', 1;\n");
            }
            return script.ToString();
        }

        private static string Predicate(ComparisonPlan plan, object[] values, Func<ColumnMapping, bool> include)
        {
            return string.Join(" AND ", plan.Columns.Select((m, i) => new { m, i }).Where(x => include(x.m))
                .Select(x => SqlText.Equal(x.m.Target, values[x.i])));
        }

        public static string RowStatement(ComparisonPlan plan, DifferenceRow row)
        {
            string table = plan.Target.QualifiedName;
            string key = Predicate(plan, row.Target ?? row.Source, c => c.IsKey);
            var script = new StringBuilder();
            if (row.Kind == DifferenceKind.OnlySource)
            {
                script.Append("IF EXISTS (SELECT 1 FROM ").Append(table).Append(" WITH (UPDLOCK, HOLDLOCK) WHERE ").Append(key)
                    .Append(") THROW 51002, 'An insert key already exists in the target. Compare again.', 1;\n");
                var writable = plan.Columns.Select((m, i) => new { m, i }).Where(x => x.m.Target.Writable).ToList();
                script.Append("INSERT INTO ").Append(table).Append(" (").Append(string.Join(", ", writable.Select(x => SqlText.Quote(x.m.Target.Name))))
                    .Append(") VALUES (").Append(string.Join(", ", writable.Select(x => SqlText.Literal(row.Source[x.i], x.m.Target)))).Append(");\n");
                script.Append("IF @@ROWCOUNT <> 1 THROW 51003, 'Insert affected an unexpected number of rows.', 1;\n");
            }
            else
            {
                string original = Predicate(plan, row.Target, c => true);
                script.Append("IF (SELECT COUNT_BIG(*) FROM ").Append(table).Append(" WITH (UPDLOCK, HOLDLOCK) WHERE ").Append(key)
                    .Append(") <> 1 THROW 51004, 'Target key is missing or no longer unique. Compare again.', 1;\n")
                    .Append("IF NOT EXISTS (SELECT 1 FROM ").Append(table).Append(" WITH (UPDLOCK, HOLDLOCK) WHERE ").Append(original)
                    .Append(") THROW 51005, 'Target data changed after comparison. Compare again.', 1;\n");
                if (row.Kind == DifferenceKind.OnlyTarget)
                    script.Append("DELETE FROM ").Append(table).Append(" WHERE ").Append(key).Append(";\n");
                else if (row.Kind == DifferenceKind.Different)
                {
                    var changes = plan.Columns.Select((m, i) => new { m, i }).Where(x => row.Changed[x.i]).ToList();
                    if (changes.Any(x => !x.m.Target.Writable || x.m.Target.Identity || x.m.IsKey))
                        throw new InvalidOperationException("A selected difference changes an identity, computed or generated column. Exclude that column and compare again, or deselect the row.");
                    if (changes.Count == 0) throw new InvalidOperationException("No writable changes in the selected row.");
                    script.Append("UPDATE ").Append(table).Append(" SET ").Append(string.Join(", ", changes.Select(x => SqlText.Quote(x.m.Target.Name) + " = " + SqlText.Literal(row.Source[x.i], x.m.Target))))
                        .Append(" WHERE ").Append(key).Append(";\n");
                }
                else throw new InvalidOperationException("Identical rows cannot be synchronized.");
                script.Append("IF @@ROWCOUNT <> 1 THROW 51006, 'Change affected an unexpected number of rows.', 1;\n");
            }
            // Catch lossy collation/type conversions and triggers that alter or suppress the intended write.
            if (row.Kind != DifferenceKind.OnlyTarget)
            {
                string expected = Predicate(plan, row.Source, c => c.Target.Writable);
                // Comparison options intentionally leave equivalent strings unchanged on updates.
                if (row.Kind == DifferenceKind.Different)
                {
                    var values = row.Target.ToArray();
                    for (int i = 0; i < values.Length; i++) if (row.Changed[i]) values[i] = row.Source[i];
                    expected = Predicate(plan, values, c => c.Target.Writable);
                }
                script.Append("IF NOT EXISTS (SELECT 1 FROM ").Append(table).Append(" WITH (UPDLOCK, HOLDLOCK) WHERE ").Append(expected)
                    .Append(") THROW 51007, 'Target did not preserve the requested values. Check conversions and triggers.', 1;\n");
            }
            else script.Append("IF EXISTS (SELECT 1 FROM ").Append(table).Append(" WITH (UPDLOCK, HOLDLOCK) WHERE ").Append(key)
                .Append(") THROW 51008, 'Target row remains after deletion. Check triggers.', 1;\n");
            return script.ToString();
        }

        public static void Write(ComparisonResult result, TextWriter writer, CancellationToken token, IProgress<ComparisonProgress> progress = null)
        {
            Validate(result);
            var plan = result.Plan;
            bool identity = NeedsIdentityInsert(result);
            writer.WriteLine("-- Axial SQL Tools - source to target data synchronization");
            writer.WriteLine("-- Generated UTC: " + DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture));
            writer.WriteLine("-- Selected rows: " + result.SelectedCount.ToString(CultureInfo.InvariantCulture));
            writer.WriteLine("-- Review before running. Constraints and triggers remain enabled. The entire batch is atomic.");
            writer.WriteLine("-- Run in a dedicated query session connected to the comparison target.");
            writer.WriteLine("SET NOCOUNT ON;\nSET XACT_ABORT ON;\nIF @@TRANCOUNT <> 0 THROW 51009, 'An existing transaction is open. Use a new session.', 1;");
            writer.Write(DestinationGuard(plan));
            writer.WriteLine("BEGIN TRY\nBEGIN TRANSACTION;");
            writer.Write(SchemaGuard(plan));
            if (identity) writer.WriteLine("SET IDENTITY_INSERT " + plan.Target.QualifiedName + " ON;");
            long count = 0;
            foreach (var row in result.SelectedRows(token))
            {
                token.ThrowIfCancellationRequested();
                writer.WriteLine(RowStatement(plan, row));
                if (++count % 100 == 0) progress?.Report(new ComparisonProgress { Phase = "Generating script", Rows = count });
            }
            token.ThrowIfCancellationRequested();
            if (identity) writer.WriteLine("SET IDENTITY_INSERT " + plan.Target.QualifiedName + " OFF;");
            writer.WriteLine("COMMIT TRANSACTION;\nEND TRY\nBEGIN CATCH\nIF XACT_STATE() <> 0 ROLLBACK TRANSACTION;");
            if (identity) writer.WriteLine("BEGIN TRY\nSET IDENTITY_INSERT " + plan.Target.QualifiedName + " OFF;\nEND TRY\nBEGIN CATCH\nPRINT 'Could not reset IDENTITY_INSERT. Close this session.';\nEND CATCH;");
            writer.WriteLine("THROW;\nEND CATCH;");
        }
    }
}
