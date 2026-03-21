using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace TestApp
{
    /// <summary>
    /// Manages a small JSON manifest file (<c>_runs_manifest.json</c>) written
    /// inside the runs folder.  The manifest records the set of .xlsx run files
    /// present at the time of the last successful trends generation.
    ///
    /// The auto-watch timer calls <see cref="HasChanged"/> each tick.  If the
    /// live file set differs from the manifest — whether a file was added OR
    /// deleted — it returns true and a human-readable description of what changed.
    /// After regeneration, the caller writes a fresh manifest via <see cref="Write"/>.
    /// </summary>
    public static class TrendsManifest
    {
        private const string ManifestFileName = "_runs_manifest.json";

        // ── Manifest model ────────────────────────────────────────────────────

        private class Manifest
        {
            public DateTime      UpdatedAt { get; set; } = DateTime.Now;
            public string        Customer  { get; set; } = "";
            /// <summary>Filenames only (no directory path), sorted.</summary>
            public List<string>  Files     { get; set; } = new();
        }

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
        /// <paramref name="runsFolder"/>, excluding the trends output file.
        /// Safe to call even if the folder doesn't exist yet.
        /// </summary>
        public static void Write(string runsFolder, string customerName)
        {
            try
            {
                var files = GetRunFiles(runsFolder, customerName);
                var manifest = new Manifest
                {
                    UpdatedAt = DateTime.Now,
                    Customer  = customerName,
                    Files     = files
                };
                string path = ManifestPath(runsFolder);
                File.WriteAllText(path,
                    JsonSerializer.Serialize(manifest,
                        new JsonSerializerOptions { WriteIndented = true }));
            }
            catch { /* non-fatal */ }
        }

        /// <summary>
        /// Compares the current folder contents against the saved manifest.
        /// Returns <c>(true, description)</c> if anything changed (add or delete),
        /// or <c>(false, "")</c> if everything matches.
        /// </summary>
        public static (bool Changed, string Description) HasChanged(
            string runsFolder, string customerName)
        {
            try
            {
                var live = GetRunFiles(runsFolder, customerName);
                var saved = ReadManifest(runsFolder, customerName);

                var added   = live.Except(saved, StringComparer.OrdinalIgnoreCase).ToList();
                var removed = saved.Except(live, StringComparer.OrdinalIgnoreCase).ToList();

                if (added.Count == 0 && removed.Count == 0)
                    return (false, "");

                var parts = new List<string>();
                if (added.Count   > 0) parts.Add($"{added.Count} file(s) added");
                if (removed.Count > 0) parts.Add($"{removed.Count} file(s) removed");
                return (true, string.Join(", ", parts));
            }
            catch
            {
                // If we can't read the manifest treat it as changed so we regenerate
                return (true, "manifest unreadable");
            }
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static string ManifestPath(string runsFolder)
            => Path.Combine(runsFolder, ManifestFileName);

        /// <summary>
        /// Returns a sorted list of .xlsx filenames present in the folder,
        /// excluding the trends output file and the manifest itself.
        /// </summary>
        private static List<string> GetRunFiles(string runsFolder, string customerName)
        {
            string trendsFile = customerName + "_Trends.xlsx";
            return Directory.GetFiles(runsFolder, "*.xlsx", SearchOption.TopDirectoryOnly)
                .Select(Path.GetFileName)
                .Where(f => f != null
                    && !f.Equals(trendsFile,        StringComparison.OrdinalIgnoreCase)
                    && !f.Equals(ManifestFileName,  StringComparison.OrdinalIgnoreCase))
                .Cast<string>()
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        /// <summary>
        /// Reads the saved manifest and returns the recorded file list.
        /// Returns an empty list if the manifest doesn't exist or can't be parsed.
        /// </summary>
        private static List<string> ReadManifest(string runsFolder, string customerName)
        {
            string path = ManifestPath(runsFolder);
            if (!File.Exists(path)) return new List<string>();

            try
            {
                var manifest = JsonSerializer.Deserialize<Manifest>(File.ReadAllText(path));
                return manifest?.Files ?? new List<string>();
            }
            catch { return new List<string>(); }
        }
    }
}
