using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace TestApp
{
    /// <summary>
    /// Manages a small JSON manifest file (<c>_runs_manifest.json</c>) written
    /// inside the runs folder.  The manifest records the set of .xlsx run files
    /// present at the time of the last successful trends generation, along with
    /// each file's size and last-write timestamp.
    ///
    /// Using size + last-write as a lightweight fingerprint means:
    ///   • A renamed file (same size + date, different name) is detected as a
    ///     rename rather than "one added + one removed", avoiding false triggers.
    ///   • A modified file (same name, different size or date) triggers
    ///     regeneration correctly.
    ///   • No SHA hashing — opening large xlsx files just to hash them would be
    ///     expensive on every watch tick.
    ///
    /// Backwards-compatible: manifests written by the old filename-only format
    /// are detected and treated as "changed" so a fresh manifest is written on
    /// the next successful generation.
    /// </summary>
    public static class TrendsManifest
    {
        private const string ManifestFileName = "_runs_manifest.json";

        // ── Manifest model ────────────────────────────────────────────────────

        /// <summary>Fingerprint for a single run file.</summary>
        private class FileEntry
        {
            public string   Name         { get; set; } = "";
            public long     SizeBytes    { get; set; } = 0;
            public DateTime LastWriteUtc { get; set; } = DateTime.MinValue;

            /// <summary>Two entries are the same physical file if their
            /// fingerprint (size + last-write) matches, regardless of name.</summary>
            public bool SameContent(FileEntry other)
                => SizeBytes == other.SizeBytes
                && Math.Abs((LastWriteUtc - other.LastWriteUtc).TotalSeconds) < 2;
        }

        private class Manifest
        {
            public int              Version    { get; set; } = 2;
            public DateTime         UpdatedAt  { get; set; } = DateTime.Now;
            public string           Customer   { get; set; } = "";
            public List<FileEntry>  Entries    { get; set; } = new();

            // Legacy field — v1 manifests stored only filenames.
            // Kept for deserialisation only; never written.
            [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
            public List<string>? Files { get; set; } = null;
        }

        private static readonly JsonSerializerOptions _jsonOpts = new()
        {
            WriteIndented = true,
            PropertyNameCaseInsensitive = true,
        };

        // ── Public API ────────────────────────────────────────────────────────

        /// <summary>
        /// Deletes the manifest file if it exists. Called when watch starts
        /// to ensure the first tick always triggers a fresh generation.
        /// </summary>
        public static void Delete(string runsFolder)
        {
            try
            {
                string path = ManifestPath(runsFolder);
                if (File.Exists(path)) File.Delete(path);
            }
            catch { }
        }

        /// <summary>
        /// Writes (or overwrites) the manifest with the current .xlsx files in
        /// <paramref name="runsFolder"/>, including size and last-write fingerprints.
        /// Safe to call even if the folder doesn't exist yet.
        /// </summary>
        public static void Write(string runsFolder, string customerName)
        {
            try
            {
                var entries = GetFileEntries(runsFolder);
                var manifest = new Manifest
                {
                    Version   = 2,
                    UpdatedAt = DateTime.Now,
                    Customer  = customerName,
                    Entries   = entries,
                };
                File.WriteAllText(ManifestPath(runsFolder),
                    JsonSerializer.Serialize(manifest, _jsonOpts));
            }
            catch { /* non-fatal */ }
        }

        /// <summary>
        /// Compares the current folder contents against the saved manifest.
        /// Returns <c>(true, description)</c> if anything changed,
        /// or <c>(false, "")</c> if everything matches.
        ///
        /// Change types detected:
        ///   • Added   — new file with no fingerprint match in the manifest
        ///   • Removed — manifest entry with no fingerprint match on disk
        ///   • Modified — same filename, different size or last-write
        ///   • Renamed — fingerprint match found under a different name
        ///     (a rename on its own does NOT trigger regeneration — the content
        ///      is the same, only the display label would change on next generate)
        /// </summary>
        public static (bool Changed, string Description) HasChanged(
            string runsFolder, string customerName)
        {
            try
            {
                var live   = GetFileEntries(runsFolder);
                var saved  = ReadManifest(runsFolder);

                // Legacy manifest (v1, filename-only) → treat as changed so a
                // fresh v2 manifest is written after the next generation.
                if (saved == null)
                    return (true, "manifest updated to new format");

                // ── Match live files against saved entries ────────────────────

                var unmatchedLive  = new List<FileEntry>(live);
                var unmatchedSaved = new List<FileEntry>(saved);

                // First pass: exact name + fingerprint match → remove from both
                foreach (var s in saved)
                {
                    var exact = unmatchedLive.FirstOrDefault(l =>
                        string.Equals(l.Name, s.Name, StringComparison.OrdinalIgnoreCase)
                        && l.SameContent(s));
                    if (exact != null)
                    {
                        unmatchedLive.Remove(exact);
                        unmatchedSaved.Remove(s);
                    }
                }

                // Second pass: fingerprint-only match → rename (content unchanged)
                var renames = new List<(string OldName, string NewName)>();
                foreach (var s in unmatchedSaved.ToList())
                {
                    var renamed = unmatchedLive.FirstOrDefault(l => l.SameContent(s));
                    if (renamed != null)
                    {
                        renames.Add((s.Name, renamed.Name));
                        unmatchedLive.Remove(renamed);
                        unmatchedSaved.Remove(s);
                    }
                }

                // What's left: truly added, truly removed, or modified (same name, diff content)
                var added    = unmatchedLive
                    .Where(l => !unmatchedSaved.Any(s =>
                        string.Equals(s.Name, l.Name, StringComparison.OrdinalIgnoreCase)))
                    .ToList();
                var removed  = unmatchedSaved
                    .Where(s => !unmatchedLive.Any(l =>
                        string.Equals(l.Name, s.Name, StringComparison.OrdinalIgnoreCase)))
                    .ToList();
                var modified = unmatchedLive
                    .Where(l => unmatchedSaved.Any(s =>
                        string.Equals(s.Name, l.Name, StringComparison.OrdinalIgnoreCase)))
                    .ToList();

                // Renames alone don't require regeneration — the run data is identical.
                // Only content changes (add / remove / modify) trigger regeneration.
                bool contentChanged = added.Count > 0 || removed.Count > 0 || modified.Count > 0;

                if (!contentChanged && renames.Count == 0)
                    return (false, "");

                if (!contentChanged && renames.Count > 0)
                    return (false, "");  // rename only — no regeneration needed

                var parts = new List<string>();
                if (added.Count    > 0) parts.Add($"{added.Count} file(s) added");
                if (removed.Count  > 0) parts.Add($"{removed.Count} file(s) removed");
                if (modified.Count > 0) parts.Add($"{modified.Count} file(s) modified");
                if (renames.Count  > 0) parts.Add($"{renames.Count} file(s) renamed");

                return (true, string.Join(", ", parts));
            }
            catch
            {
                return (true, "manifest unreadable");
            }
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static string ManifestPath(string runsFolder)
            => Path.Combine(runsFolder, ManifestFileName);

        /// <summary>
        /// Returns FileEntry records for every eligible .xlsx in the folder.
        /// </summary>
        private static List<FileEntry> GetFileEntries(string runsFolder)
        {
            try
            {
                return Directory.GetFiles(runsFolder, "*.xlsx", SearchOption.TopDirectoryOnly)
                    .Select(path => new { path, name = Path.GetFileName(path) })
                    .Where(f => f.name != null
                        && !f.name!.EndsWith("_Trends.xlsx", StringComparison.OrdinalIgnoreCase)
                        && !f.name.Equals(ManifestFileName, StringComparison.OrdinalIgnoreCase))
                    .Select(f =>
                    {
                        var fi = new FileInfo(f.path);
                        return new FileEntry
                        {
                            Name         = f.name!,
                            SizeBytes    = fi.Exists ? fi.Length : 0,
                            LastWriteUtc = fi.Exists ? fi.LastWriteTimeUtc : DateTime.MinValue,
                        };
                    })
                    .OrderBy(e => e.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            catch { return new List<FileEntry>(); }
        }

        /// <summary>
        /// Reads the saved manifest entries. Returns null if the file doesn't
        /// exist, can't be parsed, or is a legacy v1 filename-only manifest.
        /// </summary>
        private static List<FileEntry>? ReadManifest(string runsFolder)
        {
            string path = ManifestPath(runsFolder);
            if (!File.Exists(path)) return null;

            try
            {
                var json = File.ReadAllText(path);
                var manifest = JsonSerializer.Deserialize<Manifest>(json, _jsonOpts);
                if (manifest == null) return null;

                // Legacy v1: had Files list but no Entries
                if (manifest.Version < 2 || manifest.Entries.Count == 0 && manifest.Files != null)
                    return null;  // signal caller to treat as changed

                return manifest.Entries;
            }
            catch { return null; }
        }
    }
}
