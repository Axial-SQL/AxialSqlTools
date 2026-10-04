using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace AxialSqlTools.PivotGrid
{
    internal sealed class PivotSavedLayout
    {
        public string Name { get; }
        public PivotRequest Request { get; }

        internal PivotSavedLayout(string name, PivotRequest request)
        {
            Name = name;
            Request = request.Copy();
        }
    }

    /// <summary>
    /// Stores only pivot configuration. Source records are never persisted.
    /// Exact schema matching makes saved ordinal references safe for duplicate column names.
    /// </summary>
    internal static class PivotLayoutStore
    {
        private const string SettingsKey = "PivotGridLayouts";
        private const int CurrentVersion = 1;
        private const int MaxSchemas = 40;
        private const int MaxNamedPerSchema = 20;
        private const int MaxNamedTotal = 200;
        private const int MaxRequestCharacters = 65536;
        private const int MaxStoreCharacters = 2097152;
        private static readonly object Sync = new object();

        [ThreadStatic]
        private static string lastError;
        internal static string LastError => lastError;

        internal static IReadOnlyList<PivotSavedLayout> GetLayouts(PivotSnapshot snapshot)
        {
            lock (Sync)
            {
                lastError = null;
                if (!TryRead(out var store) || !TrySchema(snapshot, out var schema))
                    return new PivotSavedLayout[0];
                var entry = store.Schemas.FirstOrDefault(s => s.Schema == schema);
                if (entry == null) return new PivotSavedLayout[0];
                foreach (var layout in entry.Named)
                    if (!ValidateRequest(snapshot, layout.Request)) return new PivotSavedLayout[0];
                return entry.Named.OrderBy(l => l.Name, StringComparer.CurrentCultureIgnoreCase)
                    .Select(l => new PivotSavedLayout(l.Name, l.Request)).ToList().AsReadOnly();
            }
        }

        internal static PivotRequest GetLast(PivotSnapshot snapshot)
        {
            lock (Sync)
            {
                lastError = null;
                if (!TryRead(out var store) || !TrySchema(snapshot, out var schema)) return null;
                var request = store.Schemas.FirstOrDefault(s => s.Schema == schema)?.Last;
                return request != null && ValidateRequest(snapshot, request) ? request.Copy() : null;
            }
        }

        internal static bool SaveNamed(PivotSnapshot snapshot, string name, PivotRequest request)
        {
            lock (Sync)
            {
                lastError = null;
                name = name?.Trim();
                if (!ValidName(name)) return Fail("Enter a layout name containing 1 to 80 characters.");
                if (!ValidateRequest(snapshot, request) || !TryRead(out var store) ||
                    !TrySchema(snapshot, out var schema)) return false;
                var entry = GetOrCreate(store, schema);
                if (entry == null) return false;
                var existing = entry.Named.FirstOrDefault(l => string.Equals(l.Name, name, StringComparison.OrdinalIgnoreCase));
                if (existing == null)
                {
                    if (entry.Named.Count >= MaxNamedPerSchema)
                        return Fail("This result schema already has 20 saved layouts. Delete one before adding another.");
                    if (store.Schemas.Sum(s => s.Named.Count) >= MaxNamedTotal)
                        return Fail("There are already 200 saved pivot layouts. Delete an existing layout before adding another.");
                    entry.Named.Add(new NamedLayout { Name = name, Request = request.Copy() });
                }
                else
                {
                    existing.Name = name;
                    existing.Request = request.Copy();
                }
                entry.LastUsedUtc = DateTime.UtcNow;
                return Save(store);
            }
        }

        internal static bool SaveLast(PivotSnapshot snapshot, PivotRequest request)
        {
            lock (Sync)
            {
                lastError = null;
                if (!ValidateRequest(snapshot, request) || !TryRead(out var store) ||
                    !TrySchema(snapshot, out var schema)) return false;
                var entry = GetOrCreate(store, schema);
                if (entry == null) return false;
                entry.Last = request.Copy();
                entry.LastUsedUtc = DateTime.UtcNow;
                return Save(store);
            }
        }

        internal static bool DeleteNamed(PivotSnapshot snapshot, string name)
        {
            lock (Sync)
            {
                lastError = null;
                name = name?.Trim();
                if (!ValidName(name)) return Fail("Select a saved layout to delete.");
                if (!TryRead(out var store) || !TrySchema(snapshot, out var schema)) return false;
                var entry = store.Schemas.FirstOrDefault(s => s.Schema == schema);
                if (entry == null) return true;
                int removed = entry.Named.RemoveAll(l => string.Equals(l.Name, name, StringComparison.OrdinalIgnoreCase));
                if (removed == 0) return true;
                if (entry.Named.Count == 0 && entry.Last == null) store.Schemas.Remove(entry);
                return Save(store);
            }
        }

        private static SchemaLayouts GetOrCreate(LayoutStore store, string schema)
        {
            var entry = store.Schemas.FirstOrDefault(s => s.Schema == schema);
            if (entry != null) return entry;
            if (store.Schemas.Count >= MaxSchemas)
            {
                // Only discard remembered last layouts. A named preset requires explicit deletion.
                var oldest = store.Schemas.Where(s => s.Named.Count == 0).OrderBy(s => s.LastUsedUtc).FirstOrDefault();
                if (oldest == null)
                {
                    Fail("Saved pivot layouts cover 40 different result schemas. Delete the named layouts for an unused schema before saving another.");
                    return null;
                }
                store.Schemas.Remove(oldest);
            }
            entry = new SchemaLayouts { Schema = schema, LastUsedUtc = DateTime.UtcNow };
            store.Schemas.Add(entry);
            return entry;
        }

        private static bool TrySchema(PivotSnapshot snapshot, out string schema)
        {
            schema = null;
            if (snapshot?.Fields == null || snapshot.Fields.Any(f => f == null))
                return Fail("The result columns are unavailable. Open Pivot Grid from a completed query.");
            // Serialize separate properties rather than joining names with delimiters: SQL aliases
            // can contain any delimiter, and duplicate/unnamed fields must retain their positions.
            string signature = JsonConvert.SerializeObject(snapshot.Fields.Select(f => new
            {
                f.Index,
                f.Name,
                Type = f.DataType?.FullName
            }), Formatting.None);
            using (var hash = SHA256.Create())
                schema = BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(signature))).Replace("-", "");
            return true;
        }

        private static bool ValidateRequest(PivotSnapshot snapshot, PivotRequest request)
        {
            try
            {
                if (!WellFormedRequest(request))
                    return Fail("The pivot layout is incomplete or invalid. Choose valid fields and calculations before saving.");
                if (snapshot?.Fields == null || snapshot.Fields.Any(f => f == null))
                    return Fail("The result columns are unavailable.");
                PivotEngine.Validate(snapshot, request);
                if (JsonConvert.SerializeObject(request, Formatting.None).Length > MaxRequestCharacters)
                    return Fail("This layout has too much filter text to save. Reduce the selected values or filter criteria.");
                return true;
            }
            catch (InvalidOperationException ex)
            {
                return Fail(ex.Message);
            }
            catch (JsonException)
            {
                return Fail("The pivot layout could not be serialized. Check its fields and filter criteria.");
            }
        }

        private static bool TryRead(out LayoutStore store)
        {
            store = null;
            string json = SettingsFileStore.GetValue(SettingsKey);
            if (string.IsNullOrWhiteSpace(json))
            {
                store = new LayoutStore();
                return true;
            }
            if (json.Length > MaxStoreCharacters)
                return Fail("Saved pivot layouts exceed the supported size. Existing settings have been preserved.");
            try
            {
                store = JsonConvert.DeserializeObject<LayoutStore>(json, new JsonSerializerSettings
                {
                    TypeNameHandling = TypeNameHandling.None,
                    MaxDepth = 32
                });
                if (store == null || store.Version != CurrentVersion)
                    return Fail("Saved pivot layouts use an unsupported format. Existing settings have been preserved.");
                if (store.Schemas == null || store.Schemas.Count > MaxSchemas ||
                    store.Schemas.Any(s => s == null || string.IsNullOrEmpty(s.Schema) || s.Schema.Length != 64 ||
                        !s.Schema.All(Uri.IsHexDigit) || s.Named == null || s.Named.Count > MaxNamedPerSchema ||
                        s.Named.Any(l => l == null || !ValidName(l.Name) || !WellFormedRequest(l.Request)) ||
                        s.Named.Select(l => l.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() != s.Named.Count ||
                        (s.Last != null && !WellFormedRequest(s.Last))) ||
                    store.Schemas.Select(s => s.Schema).Distinct(StringComparer.Ordinal).Count() != store.Schemas.Count ||
                    store.Schemas.Sum(s => s.Named.Count) > MaxNamedTotal)
                    return Fail("Saved pivot layouts contain invalid data. Existing settings have been preserved.");
                return true;
            }
            catch (JsonException)
            {
                return Fail("Saved pivot layouts contain invalid JSON. Correct the PivotGridLayouts setting before saving; existing settings have been preserved.");
            }
        }

        private static bool WellFormedRequest(PivotRequest request)
        {
            return request != null && request.Rows != null && request.Columns != null &&
                request.Measures != null && request.Measures.Length > 0 && request.Filters != null &&
                request.Rows.Concat(request.Columns).All(f => f >= 0) &&
                request.Measures.All(m => m != null && m.Value >= -1 && Enum.IsDefined(typeof(PivotAggregation), m.Aggregation)) &&
                request.Filters.All(f => f != null && f.Field >= 0 && Enum.IsDefined(typeof(PivotFilterOperator), f.Operator) &&
                    (f.Operator != PivotFilterOperator.In || (f.Values != null && f.Values.Length > 0)));
        }

        private static bool Save(LayoutStore store)
        {
            try
            {
                string json = JsonConvert.SerializeObject(store, Formatting.None);
                if (json.Length > MaxStoreCharacters)
                    return Fail("Saved pivot layouts exceed the supported size. Delete unused layouts or shorten filter criteria.");
                // SettingsFileStore updates only this key and performs its usual atomic file write.
                if (SettingsFileStore.SaveValue(SettingsKey, json)) return true;
                return Fail(SettingsFileStore.LastSaveError ?? "The pivot layout could not be saved.");
            }
            catch (JsonException)
            {
                return Fail("The pivot layout could not be serialized. Existing settings have been preserved.");
            }
        }

        private static bool ValidName(string name) => !string.IsNullOrWhiteSpace(name) &&
            name.Length <= 80 && name == name.Trim() && !name.Any(char.IsControl);

        private static bool Fail(string message)
        {
            lastError = message;
            return false;
        }

        private sealed class LayoutStore
        {
            [JsonProperty(Required = Required.Always)]
            public int Version { get; set; } = CurrentVersion;
            [JsonProperty(Required = Required.Always)]
            public List<SchemaLayouts> Schemas { get; set; } = new List<SchemaLayouts>();
        }

        private sealed class SchemaLayouts
        {
            [JsonProperty(Required = Required.Always)]
            public string Schema { get; set; }
            public DateTime LastUsedUtc { get; set; }
            public PivotRequest Last { get; set; }
            [JsonProperty(Required = Required.Always)]
            public List<NamedLayout> Named { get; set; } = new List<NamedLayout>();
        }

        private sealed class NamedLayout
        {
            [JsonProperty(Required = Required.Always)]
            public string Name { get; set; }
            [JsonProperty(Required = Required.Always)]
            public PivotRequest Request { get; set; }
        }
    }
}
