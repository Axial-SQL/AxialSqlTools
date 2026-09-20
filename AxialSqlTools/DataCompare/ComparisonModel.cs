using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace AxialSqlTools.DataCompare
{
    public enum DifferenceKind { Different, OnlySource, OnlyTarget, Identical }

    public sealed class TableColumn
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public short Length { get; set; }
        public byte Precision { get; set; }
        public byte Scale { get; set; }
        public bool Nullable { get; set; }
        public bool Identity { get; set; }
        public bool Computed { get; set; }
        public bool Generated { get; set; }
        public bool Encrypted { get; set; }
        public bool HasDefault { get; set; }
        public string Collation { get; set; }
        public bool Writable => !Computed && !Generated && Type != "timestamp" && Type != "rowversion";
        public bool Supported => !Encrypted && ValueCodec.SupportedTypes.Contains(Type);
        public string Display => Name + " (" + Type + ")" + (!Supported ? " - unsupported" : !Writable ? " - read only" : "");
        public override string ToString() => Display;
    }

    public sealed class TableKey
    {
        public string Name { get; set; }
        public bool Primary { get; set; }
        public List<string> Columns { get; } = new List<string>();
    }

    public sealed class TableSchema
    {
        public string Server { get; set; }
        public string Database { get; set; }
        public string Schema { get; set; }
        public string Name { get; set; }
        public bool HistoryTable { get; set; }
        public List<TableColumn> Columns { get; } = new List<TableColumn>();
        public List<TableKey> Keys { get; } = new List<TableKey>();
        public string QualifiedName => SqlText.Quote(Schema) + "." + SqlText.Quote(Name);
        public override string ToString() => QualifiedName;
    }

    public sealed class ColumnMapping : INotifyPropertyChanged
    {
        private TableColumn target;
        private bool include;
        private bool key;
        public TableColumn Source { get; set; }
        public TableColumn Target { get => target; set { target = value; Changed(nameof(Target)); } }
        public bool Include { get => include; set { include = value; if (!value) IsKey = false; Changed(nameof(Include)); } }
        public bool IsKey { get => key; set { key = value; if (value) Include = true; Changed(nameof(IsKey)); } }
        public event PropertyChangedEventHandler PropertyChanged;
        private void Changed(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        public ColumnMapping Copy() => new ColumnMapping { Source = Source, Target = Target, Include = Include, IsKey = IsKey };
    }

    public sealed class ComparisonOptions
    {
        public bool IgnoreCase { get; set; }
        public bool IgnoreTrailingSpaces { get; set; }
        public bool UseSnapshotIsolation { get; set; }
        public int CommandTimeoutSeconds { get; set; } = 120;
        public long SortMemoryBytes { get; set; } = 32L * 1024 * 1024;
        public long MaxTemporaryBytes { get; set; } = 20L * 1024 * 1024 * 1024;
        public int MaxCellBytes { get; set; } = 16 * 1024 * 1024;
        public ComparisonOptions Copy() => (ComparisonOptions)MemberwiseClone();
    }

    public sealed class ComparisonPlan
    {
        public TableSchema Source { get; set; }
        public TableSchema Target { get; set; }
        public List<ColumnMapping> Columns { get; set; }
        public ComparisonOptions Options { get; set; } = new ComparisonOptions();
        public int[] KeyOrdinals => Columns.Select((c, i) => new { c, i }).Where(x => x.c.IsKey).Select(x => x.i).ToArray();

        public void Validate()
        {
            if (Source == null || Target == null || Columns == null || Columns.Count == 0)
                throw new InvalidOperationException("Select two tables and at least one column.");
            if (!Columns.Any(c => c.IsKey))
                throw new InvalidOperationException("Select one or more key columns that uniquely identify every row on both sides.");
            if (Columns.Count > 4096) throw new InvalidOperationException("Compare at most 4096 columns at a time.");
            if (Source.Server == Target.Server && Source.Database == Target.Database && Source.Schema == Target.Schema && Source.Name == Target.Name)
                throw new InvalidOperationException("Choose two different tables or databases.");
            if (Columns.Any(c => !c.Include || c.Source == null || c.Target == null))
                throw new InvalidOperationException("Every included source column needs a target mapping.");
            if (Columns.Select(c => c.Target.Name).Distinct(StringComparer.Ordinal).Count() != Columns.Count ||
                Columns.Select(c => c.Source.Name).Distinct(StringComparer.Ordinal).Count() != Columns.Count)
                throw new InvalidOperationException("A column can only be mapped once.");
            foreach (var column in Columns)
            {
                if (!column.Source.Supported || !column.Target.Supported)
                    throw new InvalidOperationException("Unsupported or encrypted column: " + column.Source.Name + ". Exclude it to continue.");
                if (column.Source.Type != column.Target.Type || column.Source.Scale != column.Target.Scale ||
                    column.Source.Precision != column.Target.Precision || column.Source.Length != column.Target.Length)
                    throw new InvalidOperationException("Mapped types, lengths, precision and scale must match: " + column.Source.Name + ".");
                if (column.IsKey && (!column.Source.Writable || !column.Target.Writable))
                    throw new InvalidOperationException("Computed, generated and rowversion columns cannot be comparison keys.");
            }
            if (Options.CommandTimeoutSeconds < 1 || Options.SortMemoryBytes < 1024 || Options.MaxTemporaryBytes < 1024 || Options.MaxCellBytes < 1)
                throw new InvalidOperationException("Timeout and storage limits must be positive.");
        }

        public static List<ColumnMapping> AutoMap(TableSchema source, TableSchema target)
        {
            var result = source.Columns.Select(c =>
            {
                var match = target.Columns.FirstOrDefault(t => t.Name == c.Name);
                if (match == null)
                {
                    var candidates = target.Columns.Where(t => string.Equals(t.Name, c.Name, StringComparison.OrdinalIgnoreCase)).ToList();
                    if (candidates.Count == 1) match = candidates[0];
                }
                return new ColumnMapping { Source = c, Target = match,
                    Include = match != null && c.Supported && match.Supported && c.Writable && match.Writable };
            }).ToList();
            var key = source.Keys.OrderByDescending(k => k.Primary).ThenBy(k => k.Columns.Count)
                .FirstOrDefault(k => k.Columns.All(n => result.Any(m => m.Source.Name == n && m.Include)));
            if (key != null)
                foreach (var mapping in result) mapping.IsKey = key.Columns.Contains(mapping.Source.Name);
            return result;
        }
    }

    public sealed class DifferenceRow
    {
        public long Index { get; set; }
        public DifferenceKind Kind { get; set; }
        public object[] Source { get; set; }
        public object[] Target { get; set; }
        public bool[] Changed { get; set; }
    }

    public sealed class ComparisonProgress
    {
        public string Phase { get; set; }
        public long Rows { get; set; }
        public override string ToString() => Phase + ": " + Rows.ToString("N0") + " rows";
    }
}
