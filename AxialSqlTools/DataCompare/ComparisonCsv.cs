using System;
using System.IO;
using System.Linq;
using System.Threading;

namespace AxialSqlTools.DataCompare
{
    public static class ComparisonCsv
    {
        public static void Write(ComparisonResult result, TextWriter writer, CancellationToken token, IProgress<ComparisonProgress> progress = null)
        {
            writer.WriteLine("Category,Selected,Column,IsKey,Changed,SourceRowExists,SourceIsNull,SourceValue,TargetRowExists,TargetIsNull,TargetValue,CategoryRowIndex");
            long count = 0;
            foreach (DifferenceKind kind in Enum.GetValues(typeof(DifferenceKind)))
                foreach (var row in result.Read(kind))
                {
                    token.ThrowIfCancellationRequested();
                    for (int i = 0; i < result.Plan.Columns.Count; i++)
                    {
                        var mapping = result.Plan.Columns[i];
                        writer.WriteLine(string.Join(",", new[] { kind.ToString(), result.Selection.IsSelected(kind, row.Index).ToString(),
                            mapping.Source.Name + " -> " + mapping.Target.Name, mapping.IsKey.ToString(), row.Changed[i].ToString(),
                            (row.Source != null).ToString(), (row.Source != null && ValueCodec.IsNull(row.Source[i])).ToString(),
                            row.Source == null || ValueCodec.IsNull(row.Source[i]) ? "" : ValueCodec.Display(row.Source[i], int.MaxValue),
                            (row.Target != null).ToString(), (row.Target != null && ValueCodec.IsNull(row.Target[i])).ToString(),
                            row.Target == null || ValueCodec.IsNull(row.Target[i]) ? "" : ValueCodec.Display(row.Target[i], int.MaxValue),
                            (row.Index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) }.Select(Escape)));
                    }
                    if (++count % 100 == 0) progress?.Report(new ComparisonProgress { Phase = "Exporting CSV", Rows = count });
                }
        }

        private static string Escape(string value)
        {
            // CSV is a report: neutralize spreadsheet formulas while retaining raw values in SQL export.
            if (!string.IsNullOrEmpty(value) && "=+-@\t\r\n".IndexOf(value[0]) >= 0) value = "'" + value;
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }
}
