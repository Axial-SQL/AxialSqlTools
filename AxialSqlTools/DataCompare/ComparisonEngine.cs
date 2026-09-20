using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace AxialSqlTools.DataCompare
{
    public static class ComparisonEngine
    {
        public static ComparisonResult Compare(ComparisonPlan plan, IEnumerable<object[]> source,
            IEnumerable<object[]> target, CancellationToken token, IProgress<ComparisonProgress> progress = null)
        {
            plan.Validate();
            var workspace = new ComparisonWorkspace(plan.Options.MaxTemporaryBytes);
            ComparisonResult result = null;
            try
            {
                var comparer = new KeyComparer(plan.KeyOrdinals);
                var sort = new ExternalRowSort(workspace, comparer, plan.Options);
                var sourceFiles = sort.Sort(source, plan.Columns.Count, token, n => Report(progress, "Reading source", n));
                var targetFiles = sort.Sort(target, plan.Columns.Count, token, n => Report(progress, "Reading target", n));
                result = new ComparisonResult(workspace, plan);
                using (var left = Unique(sort.Merge(sourceFiles, token), comparer, "source").GetEnumerator())
                using (var right = Unique(sort.Merge(targetFiles, token), comparer, "target").GetEnumerator())
                {
                    bool hasLeft = left.MoveNext(), hasRight = right.MoveNext();
                    long count = 0;
                    while (hasLeft || hasRight)
                    {
                        token.ThrowIfCancellationRequested();
                        int order = !hasLeft ? 1 : !hasRight ? -1 : comparer.Compare(left.Current, right.Current);
                        if (order < 0) { result.Add(DifferenceKind.OnlySource, left.Current, null); hasLeft = left.MoveNext(); }
                        else if (order > 0) { result.Add(DifferenceKind.OnlyTarget, null, right.Current); hasRight = right.MoveNext(); }
                        else
                        {
                            bool different = ChangedColumns(plan, left.Current, right.Current).Any(c => c);
                            result.Add(different ? DifferenceKind.Different : DifferenceKind.Identical, left.Current, right.Current);
                            hasLeft = left.MoveNext(); hasRight = right.MoveNext();
                        }
                        if (++count % 2000 == 0) Report(progress, "Comparing", count);
                    }
                    Report(progress, "Complete", count);
                }
                foreach (string file in sourceFiles.Concat(targetFiles)) workspace.Delete(file);
                token.ThrowIfCancellationRequested();
                result.Complete();
                return result;
            }
            catch
            {
                if (result != null) result.Dispose(); else workspace.Dispose();
                throw;
            }
        }

        private static IEnumerable<object[]> Unique(IEnumerable<object[]> rows, KeyComparer comparer, string side)
        {
            object[] previous = null;
            foreach (var row in rows)
            {
                if (previous != null && comparer.Compare(previous, row) == 0)
                    throw new InvalidOperationException("Duplicate comparison key in the " + side + " table. Choose a unique key. No synchronization result was created.");
                previous = row;
                yield return row;
            }
        }

        public static bool[] ChangedColumns(ComparisonPlan plan, object[] source, object[] target)
        {
            var changed = new bool[plan.Columns.Count];
            for (int i = 0; i < changed.Length; i++)
                changed[i] = source == null || target == null || ValueCodec.Compare(source[i], target[i],
                    !plan.Columns[i].IsKey && plan.Options.IgnoreCase, !plan.Columns[i].IsKey && plan.Options.IgnoreTrailingSpaces) != 0;
            return changed;
        }

        private static void Report(IProgress<ComparisonProgress> progress, string phase, long rows) =>
            progress?.Report(new ComparisonProgress { Phase = phase, Rows = rows });
    }
}
