using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace AxialSqlTools.DataCompare
{
    internal sealed class ComparisonWorkspace : IDisposable
    {
        private readonly string directory;
        private readonly long limit;
        private long used;
        private int next;
        private readonly Dictionary<string, long> sizes = new Dictionary<string, long>();
        public ComparisonWorkspace(long limit)
        {
            this.limit = limit;
            directory = Path.Combine(Path.GetTempPath(), "AxialSQL", "DataCompare", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
        }
        public string CreatePath() => Path.Combine(directory, (++next).ToString() + ".bin");
        public void Account(string path, long length)
        {
            sizes.TryGetValue(path, out long previous);
            used += length - previous;
            sizes[path] = length;
            if (used > limit) throw new IOException("Comparison temporary storage limit exceeded. Increase the limit and compare again.");
        }
        public void Delete(string path)
        {
            File.Delete(path);
            if (sizes.TryGetValue(path, out long size)) { used -= size; sizes.Remove(path); }
        }
        public void Dispose()
        {
            try { Directory.Delete(directory, true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    internal sealed class ExternalRowSort
    {
        private const int FanIn = 16;
        private readonly ComparisonWorkspace workspace;
        private readonly IComparer<object[]> comparer;
        private readonly ComparisonOptions options;
        public ExternalRowSort(ComparisonWorkspace workspace, IComparer<object[]> comparer, ComparisonOptions options)
        { this.workspace = workspace; this.comparer = comparer; this.options = options; }

        public List<string> Sort(IEnumerable<object[]> input, int columnCount, CancellationToken token, Action<long> progress)
        {
            var runs = new List<string>();
            var rows = new List<object[]>();
            long memory = 0, count = 0;
            foreach (var row in input)
            {
                token.ThrowIfCancellationRequested();
                if (row == null || row.Length != columnCount) throw new InvalidDataException("Unexpected comparison column count.");
                long rowBytes = ValueCodec.EstimateBytes(row);
                if (rowBytes > 32L * 1024 * 1024) throw new InvalidDataException("A comparison row exceeds 32 MiB. Exclude large columns and compare again.");
                foreach (var value in row)
                    if ((value is string text && text.Length * 2L > options.MaxCellBytes) ||
                        (value is byte[] bytes && bytes.Length > options.MaxCellBytes))
                        throw new InvalidDataException("A value exceeds the comparison cell size limit. Exclude that column and compare again.");
                rows.Add(row); memory += rowBytes; count++;
                if (memory >= options.SortMemoryBytes || rows.Count >= 50000)
                {
                    runs.Add(WriteRun(rows, token)); rows.Clear(); memory = 0;
                }
                if (count % 2000 == 0) progress(count);
            }
            if (rows.Count > 0) runs.Add(WriteRun(rows, token));
            progress(count);
            while (runs.Count > FanIn)
            {
                var next = new List<string>();
                for (int i = 0; i < runs.Count; i += FanIn)
                {
                    token.ThrowIfCancellationRequested();
                    var group = runs.Skip(i).Take(FanIn).ToList();
                    if (group.Count == 1) { next.Add(group[0]); continue; }
                    next.Add(WriteRows(Merge(group, token), token));
                    foreach (string run in group) workspace.Delete(run);
                }
                runs = next;
            }
            return runs;
        }

        private string WriteRun(List<object[]> rows, CancellationToken token)
        {
            rows.Sort(comparer);
            return WriteRows(rows, token);
        }

        private string WriteRows(IEnumerable<object[]> rows, CancellationToken token)
        {
            string path = workspace.CreatePath();
            using (var writer = new BinaryWriter(File.Create(path)))
                foreach (var row in rows)
                {
                    token.ThrowIfCancellationRequested();
                    ValueCodec.WriteRow(writer, row);
                    workspace.Account(path, writer.BaseStream.Position);
                }
            return path;
        }

        private sealed class Cursor : IDisposable
        {
            public int Id;
            public object[] Row;
            public BinaryReader Reader;
            public bool Advance()
            {
                if (Reader.BaseStream.Position == Reader.BaseStream.Length) return false;
                Row = ValueCodec.ReadRow(Reader); return true;
            }
            public void Dispose() => Reader?.Dispose();
        }

        public IEnumerable<object[]> Merge(List<string> runs, CancellationToken token)
        {
            var cursors = new List<Cursor>();
            var queue = new SortedSet<Cursor>(Comparer<Cursor>.Create((a, b) =>
            {
                int order = comparer.Compare(a.Row, b.Row);
                return order == 0 ? a.Id.CompareTo(b.Id) : order;
            }));
            try
            {
                foreach (string path in runs)
                {
                    token.ThrowIfCancellationRequested();
                    var cursor = new Cursor { Id = cursors.Count, Reader = new BinaryReader(File.OpenRead(path)) };
                    cursors.Add(cursor);
                    if (cursor.Advance()) queue.Add(cursor);
                }
                while (queue.Count > 0)
                {
                    token.ThrowIfCancellationRequested();
                    var cursor = queue.Min;
                    queue.Remove(cursor);
                    yield return cursor.Row;
                    if (cursor.Advance()) queue.Add(cursor);
                }
            }
            finally { foreach (var cursor in cursors) cursor.Dispose(); }
        }
    }

    public sealed class ResultSelection
    {
        private readonly bool[] defaults = { true, true, false, false };
        private readonly HashSet<long>[] exceptions = Enumerable.Range(0, 4).Select(_ => new HashSet<long>()).ToArray();
        public bool IsSelected(DifferenceKind kind, long index) => kind != DifferenceKind.Identical && (defaults[(int)kind] ^ exceptions[(int)kind].Contains(index));
        public void Set(DifferenceKind kind, long index, bool value)
        {
            if (kind == DifferenceKind.Identical) return;
            if (value == defaults[(int)kind]) exceptions[(int)kind].Remove(index);
            else exceptions[(int)kind].Add(index);
        }
        public void SetAll(DifferenceKind kind, bool value) { defaults[(int)kind] = kind != DifferenceKind.Identical && value; exceptions[(int)kind].Clear(); }
        public long Count(DifferenceKind kind, long total) => defaults[(int)kind] ? total - exceptions[(int)kind].Count : exceptions[(int)kind].Count;
    }

    public sealed class ComparisonResult : IDisposable
    {
        private readonly ComparisonWorkspace workspace;
        private readonly string dataPath;
        private readonly string[] indexPaths;
        private BinaryWriter data;
        private BinaryWriter[] indices;
        private readonly long[] counts = new long[4];
        private bool disposed;
        public ComparisonPlan Plan { get; }
        public ResultSelection Selection { get; } = new ResultSelection();
        public DateTime CreatedUtc { get; } = DateTime.UtcNow;
        public long Count(DifferenceKind kind) => counts[(int)kind];
        public long SelectedCount => Enum.GetValues(typeof(DifferenceKind)).Cast<DifferenceKind>().Sum(k => Selection.Count(k, Count(k)));

        internal ComparisonResult(ComparisonWorkspace workspace, ComparisonPlan plan)
        {
            this.workspace = workspace; Plan = plan;
            dataPath = workspace.CreatePath();
            indexPaths = Enumerable.Range(0, 4).Select(_ => workspace.CreatePath()).ToArray();
            try
            {
                data = new BinaryWriter(File.Create(dataPath));
                indices = new BinaryWriter[4];
                for (int i = 0; i < 4; i++) indices[i] = new BinaryWriter(File.Create(indexPaths[i]));
            }
            catch { Dispose(); throw; }
        }

        internal void Add(DifferenceKind kind, object[] source, object[] target)
        {
            int group = (int)kind;
            indices[group].Write(data.BaseStream.Position);
            ValueCodec.WriteRow(data, source);
            // Keep the target's original values even when options ignore textual differences.
            ValueCodec.WriteRow(data, target);
            counts[group]++;
            workspace.Account(dataPath, data.BaseStream.Position);
            workspace.Account(indexPaths[group], indices[group].BaseStream.Position);
        }

        internal void Complete()
        {
            data.Dispose(); data = null;
            foreach (var index in indices) index.Dispose();
            indices = null;
        }

        public IEnumerable<DifferenceRow> Read(DifferenceKind kind, long start = 0, long limit = long.MaxValue)
        {
            if (disposed || data != null) throw new InvalidOperationException("Comparison results are not available.");
            if (start < 0 || limit < 0) throw new ArgumentOutOfRangeException(nameof(start));
            using (var index = new BinaryReader(File.OpenRead(indexPaths[(int)kind])))
            using (var reader = new BinaryReader(File.OpenRead(dataPath)))
            {
                if (start >= Count(kind)) yield break;
                index.BaseStream.Position = start * sizeof(long);
                long end = start + Math.Min(limit, Count(kind) - start);
                for (long i = start; i < end; i++)
                {
                    reader.BaseStream.Position = index.ReadInt64();
                    var source = ValueCodec.ReadRow(reader); var target = ValueCodec.ReadRow(reader);
                    yield return new DifferenceRow { Index = i, Kind = kind, Source = source, Target = target,
                        Changed = ComparisonEngine.ChangedColumns(Plan, source, target) };
                }
            }
        }

        public IEnumerable<DifferenceRow> SelectedRows(CancellationToken token)
        {
            foreach (var kind in new[] { DifferenceKind.OnlyTarget, DifferenceKind.Different, DifferenceKind.OnlySource })
            {
                if (Selection.Count(kind, Count(kind)) == 0) continue;
                foreach (var row in Read(kind))
                {
                    token.ThrowIfCancellationRequested();
                    if (Selection.IsSelected(kind, row.Index)) yield return row;
                }
            }
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            data?.Dispose();
            if (indices != null) foreach (var index in indices) index?.Dispose();
            workspace.Dispose();
        }
    }
}
