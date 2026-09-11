using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Globalization;
using System.IO;
using System.Linq;

namespace AxialSqlTools.DataCompare
{
    // Explicit, version-local encoding. Never use BinaryFormatter for comparison files.
    public static class ValueCodec
    {
        public static readonly HashSet<string> SupportedTypes = new HashSet<string>(StringComparer.Ordinal)
        {
            "bigint", "int", "smallint", "tinyint", "bit", "decimal", "numeric", "money", "smallmoney",
            "float", "real", "date", "datetime", "datetime2", "smalldatetime", "datetimeoffset", "time",
            "uniqueidentifier", "char", "varchar", "nchar", "nvarchar", "text", "ntext", "xml",
            "binary", "varbinary", "image", "timestamp", "rowversion"
        };

        private enum Tag : byte { Null, String, Bytes, Int64, Int32, Int16, Byte, Boolean, Decimal, SqlDecimal, Double, Single, DateTime, Offset, Time, Guid }

        public static void WriteRow(BinaryWriter writer, object[] row)
        {
            writer.Write(row == null ? -1 : row.Length);
            if (row != null) foreach (object value in row) Write(writer, value);
        }

        public static object[] ReadRow(BinaryReader reader)
        {
            int count = reader.ReadInt32();
            if (count == -1) return null;
            if (count < 0 || count > 4096) throw new InvalidDataException("Invalid comparison row.");
            var row = new object[count];
            for (int i = 0; i < count; i++) row[i] = Read(reader);
            return row;
        }

        public static long EstimateBytes(object[] row)
        {
            long size = 64 + row.Length * 32L;
            foreach (var value in row)
            {
                if (value is string text) size += text.Length * 2L + 32;
                if (value is byte[] bytes) size += bytes.Length + 32;
            }
            return size;
        }

        public static bool IsNull(object value) => value == null || value == DBNull.Value || (value is INullable sql && sql.IsNull);

        private static void Write(BinaryWriter writer, object value)
        {
            if (IsNull(value)) { writer.Write((byte)Tag.Null); return; }
            if (value is string text)
            {
                writer.Write((byte)Tag.String);
                // Preserve UTF-16 code units, including unpaired surrogates present in SQL data.
                var bytes = new byte[text.Length * 2];
                Buffer.BlockCopy(text.ToCharArray(), 0, bytes, 0, bytes.Length);
                writer.Write(bytes.Length); writer.Write(bytes); return;
            }
            if (value is byte[] binary) { writer.Write((byte)Tag.Bytes); writer.Write(binary.Length); writer.Write(binary); return; }
            if (value is long i64) { writer.Write((byte)Tag.Int64); writer.Write(i64); return; }
            if (value is int i32) { writer.Write((byte)Tag.Int32); writer.Write(i32); return; }
            if (value is short i16) { writer.Write((byte)Tag.Int16); writer.Write(i16); return; }
            if (value is byte u8) { writer.Write((byte)Tag.Byte); writer.Write(u8); return; }
            if (value is bool boolean) { writer.Write((byte)Tag.Boolean); writer.Write(boolean); return; }
            if (value is decimal number) { writer.Write((byte)Tag.Decimal); writer.Write(number); return; }
            if (value is SqlDecimal sqlDecimal)
            {
                writer.Write((byte)Tag.SqlDecimal); writer.Write(sqlDecimal.Precision); writer.Write(sqlDecimal.Scale);
                writer.Write(sqlDecimal.IsPositive); foreach (int part in sqlDecimal.Data) writer.Write(part); return;
            }
            if (value is double f64) { writer.Write((byte)Tag.Double); writer.Write(f64); return; }
            if (value is float f32) { writer.Write((byte)Tag.Single); writer.Write(f32); return; }
            if (value is DateTime date) { writer.Write((byte)Tag.DateTime); writer.Write(date.Ticks); return; }
            if (value is DateTimeOffset offset) { writer.Write((byte)Tag.Offset); writer.Write(offset.Ticks); writer.Write(offset.Offset.Ticks); return; }
            if (value is TimeSpan time) { writer.Write((byte)Tag.Time); writer.Write(time.Ticks); return; }
            if (value is Guid guid) { writer.Write((byte)Tag.Guid); writer.Write(guid.ToByteArray()); return; }
            throw new NotSupportedException("Unsupported SQL value: " + value.GetType().Name);
        }

        private static object Read(BinaryReader reader)
        {
            switch ((Tag)reader.ReadByte())
            {
                case Tag.Null: return null;
                case Tag.String:
                    var bytes = ReadBytes(reader, reader.ReadInt32());
                    if (bytes.Length % 2 != 0) throw new InvalidDataException("Invalid string length.");
                    var chars = new char[bytes.Length / 2]; Buffer.BlockCopy(bytes, 0, chars, 0, bytes.Length); return new string(chars);
                case Tag.Bytes: return ReadBytes(reader, reader.ReadInt32());
                case Tag.Int64: return reader.ReadInt64();
                case Tag.Int32: return reader.ReadInt32();
                case Tag.Int16: return reader.ReadInt16();
                case Tag.Byte: return reader.ReadByte();
                case Tag.Boolean: return reader.ReadBoolean();
                case Tag.Decimal: return reader.ReadDecimal();
                case Tag.SqlDecimal:
                    byte precision = reader.ReadByte(), scale = reader.ReadByte(); bool positive = reader.ReadBoolean();
                    return new SqlDecimal(precision, scale, positive, new[] { reader.ReadInt32(), reader.ReadInt32(), reader.ReadInt32(), reader.ReadInt32() });
                case Tag.Double: return reader.ReadDouble();
                case Tag.Single: return reader.ReadSingle();
                case Tag.DateTime: return new DateTime(reader.ReadInt64(), DateTimeKind.Unspecified);
                case Tag.Offset: return new DateTimeOffset(reader.ReadInt64(), new TimeSpan(reader.ReadInt64()));
                case Tag.Time: return new TimeSpan(reader.ReadInt64());
                case Tag.Guid: return new Guid(ReadBytes(reader, 16));
                default: throw new InvalidDataException("Invalid comparison value tag.");
            }
        }

        private static byte[] ReadBytes(BinaryReader reader, int count)
        {
            if (count < 0 || count > 256 * 1024 * 1024) throw new InvalidDataException("Invalid comparison value length.");
            var bytes = reader.ReadBytes(count);
            if (bytes.Length != count) throw new EndOfStreamException("Comparison file is incomplete.");
            return bytes;
        }

        public static int Compare(object left, object right, bool ignoreCase = false, bool trimSpaces = false)
        {
            bool leftNull = IsNull(left), rightNull = IsNull(right);
            if (leftNull || rightNull) return leftNull == rightNull ? 0 : leftNull ? -1 : 1;
            if (left is string a && right is string b)
            {
                if (trimSpaces) { a = a.TrimEnd(' '); b = b.TrimEnd(' '); }
                return (ignoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal).Compare(a, b);
            }
            if (left is byte[] x && right is byte[] y)
            {
                for (int i = 0; i < Math.Min(x.Length, y.Length); i++)
                    if (x[i] != y[i]) return x[i].CompareTo(y[i]);
                return x.Length.CompareTo(y.Length);
            }
            if (left.GetType() != right.GetType()) throw new InvalidOperationException("Mapped columns returned incompatible value types.");
            return ((IComparable)left).CompareTo(right);
        }

        public static string Display(object value, int maxLength = 4000)
        {
            if (IsNull(value)) return "<NULL>";
            string text;
            if (value is byte[] bytes) text = "0x" + BitConverter.ToString(bytes.Take(maxLength / 2).ToArray()).Replace("-", "") + (bytes.Length > maxLength / 2 ? "..." : "");
            else if (value is DateTime date) text = date.ToString("yyyy-MM-dd HH:mm:ss.fffffff", CultureInfo.InvariantCulture);
            else if (value is DateTimeOffset offset) text = offset.ToString("o", CultureInfo.InvariantCulture);
            else if (value is TimeSpan time) text = time.ToString("c", CultureInfo.InvariantCulture);
            else text = Convert.ToString(value, CultureInfo.InvariantCulture);
            return text.Length > maxLength ? text.Substring(0, maxLength) + "..." : text;
        }
    }

    internal sealed class KeyComparer : IComparer<object[]>
    {
        private readonly int[] ordinals;
        public KeyComparer(int[] ordinals) { this.ordinals = ordinals; }
        public int Compare(object[] left, object[] right)
        {
            foreach (int ordinal in ordinals)
            {
                int result = ValueCodec.Compare(left[ordinal], right[ordinal]);
                if (result != 0) return result;
            }
            return 0;
        }
    }
}
