using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace TestApp
{
    // ── Schema version constants ──────────────────────────────────────────────
    // Bump CurrentVersion whenever fields are added/removed from a persisted class.
    // MigrateScriptLibrary / MigrateTrendsLibrary handle upgrading old files.

    public static class AppDataManager
    {
        public const int CurrentScriptLibraryVersion  = 1;
        public const int CurrentTrendsLibraryVersion  = 1;
        public const int CurrentSettingsVersion       = 1;

        private static readonly string AppDataDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "PerformanceTestUtilities");

        public static readonly string ScriptLibraryPath  = Path.Combine(AppDataDir, "script_library.json");
        public static readonly string TrendsLibraryPath  = Path.Combine(AppDataDir, "trends_library.json");
        public static readonly string SettingsPath       = Path.Combine(AppDataDir, "settings.json");

        private static readonly JsonSerializerOptions Opts = new()
        {
            WriteIndented        = true,
            DefaultIgnoreCondition = JsonIgnoreCondition.Never,
            PropertyNameCaseInsensitive = true,
        };

        // ── Ensure directory ──────────────────────────────────────────────────

        private static void EnsureDir() => Directory.CreateDirectory(AppDataDir);

        // ─────────────────────────────────────────────────────────────────────
        // Script Library
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>Versioned wrapper stored on disk for the script library.</summary>
        public class ScriptLibraryFile
        {
            public int Version { get; set; } = CurrentScriptLibraryVersion;
            public List<MainWindow.ScriptEntry> Entries { get; set; } = new();
        }

        public static List<MainWindow.ScriptEntry> LoadScriptLibrary()
        {
            try
            {
                if (!File.Exists(ScriptLibraryPath)) return new();

                var json = File.ReadAllText(ScriptLibraryPath);

                // Legacy format: bare array (version 0 — no wrapper object)
                if (json.TrimStart().StartsWith("["))
                {
                    var legacy = JsonSerializer.Deserialize<List<MainWindow.ScriptEntry>>(json, Opts) ?? new();
                    return MigrateScriptLibrary(legacy, fromVersion: 0);
                }

                var file = JsonSerializer.Deserialize<ScriptLibraryFile>(json, Opts);
                if (file == null) return new();

                return MigrateScriptLibrary(file.Entries, file.Version);
            }
            catch { return new(); }
        }

        public static void SaveScriptLibrary(List<MainWindow.ScriptEntry> entries)
        {
            try
            {
                EnsureDir();
                var file = new ScriptLibraryFile { Version = CurrentScriptLibraryVersion, Entries = entries };
                File.WriteAllText(ScriptLibraryPath, JsonSerializer.Serialize(file, Opts));
            }
            catch { }
        }

        /// <summary>Export script library to a user-chosen file.</summary>
        public static bool ExportScriptLibrary(List<MainWindow.ScriptEntry> entries, string destPath)
        {
            try
            {
                var file = new ScriptLibraryFile { Version = CurrentScriptLibraryVersion, Entries = entries };
                File.WriteAllText(destPath, JsonSerializer.Serialize(file, Opts));
                return true;
            }
            catch { return false; }
        }

        /// <summary>
        /// Import script library from a user-chosen file.
        /// Returns the imported entries, or null on failure.
        /// Existing entries with the same Id are skipped (merge, not replace).
        /// </summary>
        public static (List<MainWindow.ScriptEntry>? Imported, string? Error) ImportScriptLibrary(
            string sourcePath, List<MainWindow.ScriptEntry> existing)
        {
            try
            {
                var json = File.ReadAllText(sourcePath);
                List<MainWindow.ScriptEntry> incoming;

                if (json.TrimStart().StartsWith("["))
                    incoming = JsonSerializer.Deserialize<List<MainWindow.ScriptEntry>>(json, Opts) ?? new();
                else
                {
                    var file = JsonSerializer.Deserialize<ScriptLibraryFile>(json, Opts);
                    if (file == null) return (null, "File could not be parsed.");
                    incoming = MigrateScriptLibrary(file.Entries, file.Version);
                }

                var existingIds = new HashSet<string>(
                    System.Linq.Enumerable.Select(existing, e => e.Id),
                    StringComparer.OrdinalIgnoreCase);

                var newEntries = new List<MainWindow.ScriptEntry>();
                foreach (var e in incoming)
                    if (!existingIds.Contains(e.Id))
                        newEntries.Add(e);

                return (newEntries, null);
            }
            catch (Exception ex) { return (null, ex.Message); }
        }

        private static List<MainWindow.ScriptEntry> MigrateScriptLibrary(
            List<MainWindow.ScriptEntry> entries, int fromVersion)
        {
            // fromVersion 0 → 1: nothing structural to change, just re-save with wrapper.
            // Add future migration steps here as: if (fromVersion < N) { ... }
            return entries;
        }

        // ─────────────────────────────────────────────────────────────────────
        // Trends Library
        // ─────────────────────────────────────────────────────────────────────

        public class TrendsLibraryFile
        {
            public int Version { get; set; } = CurrentTrendsLibraryVersion;
            public List<TrendsCustomerDto> Entries { get; set; } = new();
        }

        /// <summary>DTO mirrors TrendsCustomer (which is private in MainWindow).</summary>
        public class TrendsCustomerDto
        {
            public string    Id             { get; set; } = Guid.NewGuid().ToString("N")[..8];
            public string    Name           { get; set; } = "";
            public string    RunsFolder     { get; set; } = "";
            public string    ReportsFolder  { get; set; } = "";
            public DateTime? LastGenerated  { get; set; } = null;
            public string?   LastOutput     { get; set; } = null;
        }

        public static List<TrendsCustomerDto> LoadTrendsLibrary()
        {
            try
            {
                if (!File.Exists(TrendsLibraryPath)) return new();
                var json = File.ReadAllText(TrendsLibraryPath);

                if (json.TrimStart().StartsWith("["))
                {
                    var legacy = JsonSerializer.Deserialize<List<TrendsCustomerDto>>(json, Opts) ?? new();
                    return MigrateTrendsLibrary(legacy, 0);
                }

                var file = JsonSerializer.Deserialize<TrendsLibraryFile>(json, Opts);
                if (file == null) return new();
                return MigrateTrendsLibrary(file.Entries, file.Version);
            }
            catch { return new(); }
        }

        public static void SaveTrendsLibrary(List<TrendsCustomerDto> entries)
        {
            try
            {
                EnsureDir();
                var file = new TrendsLibraryFile { Version = CurrentTrendsLibraryVersion, Entries = entries };
                File.WriteAllText(TrendsLibraryPath, JsonSerializer.Serialize(file, Opts));
            }
            catch { }
        }

        public static bool ExportTrendsLibrary(List<TrendsCustomerDto> entries, string destPath)
        {
            try
            {
                var file = new TrendsLibraryFile { Version = CurrentTrendsLibraryVersion, Entries = entries };
                File.WriteAllText(destPath, JsonSerializer.Serialize(file, Opts));
                return true;
            }
            catch { return false; }
        }

        public static (List<TrendsCustomerDto>? Imported, string? Error) ImportTrendsLibrary(
            string sourcePath, List<TrendsCustomerDto> existing)
        {
            try
            {
                var json = File.ReadAllText(sourcePath);
                List<TrendsCustomerDto> incoming;

                if (json.TrimStart().StartsWith("["))
                    incoming = JsonSerializer.Deserialize<List<TrendsCustomerDto>>(json, Opts) ?? new();
                else
                {
                    var file = JsonSerializer.Deserialize<TrendsLibraryFile>(json, Opts);
                    if (file == null) return (null, "File could not be parsed.");
                    incoming = MigrateTrendsLibrary(file.Entries, file.Version);
                }

                var existingNames = new HashSet<string>(
                    System.Linq.Enumerable.Select(existing, e => e.Name),
                    StringComparer.OrdinalIgnoreCase);

                var newEntries = new List<TrendsCustomerDto>();
                foreach (var e in incoming)
                    if (!existingNames.Contains(e.Name))
                        newEntries.Add(e);

                return (newEntries, null);
            }
            catch (Exception ex) { return (null, ex.Message); }
        }

        private static List<TrendsCustomerDto> MigrateTrendsLibrary(
            List<TrendsCustomerDto> entries, int fromVersion)
        {
            // Add migration steps here as needed.
            return entries;
        }

        // ─────────────────────────────────────────────────────────────────────
        // UI Settings
        // ─────────────────────────────────────────────────────────────────────

        public class AppSettings
        {
            public int Version { get; set; } = CurrentSettingsVersion;

            // ── Convert Response Times ───────────────────────────────────────
            public bool ConvertClubOutput    { get; set; } = false;
            public bool ConvertIncludeCharts { get; set; } = true;

            // ── JTL File Processing ──────────────────────────────────────────
            public bool JtlClubOutput        { get; set; } = false;
            public bool JtlIncludeCharts     { get; set; } = true;

            // ── BLG Conversion ───────────────────────────────────────────────
            public string BlgServerType      { get; set; } = "App";   // "App" | "Db"
            public bool   BlgProduceGraphs   { get; set; } = false;

            // ── Run Comparison ───────────────────────────────────────────────
            public string CmpMode            { get; set; } = "AllVsBaseline"; // "AllVsBaseline" | "Sequential"
            public string CmpSlaMs           { get; set; } = "";

            // ── Test Run Trends ──────────────────────────────────────────────
            public string TrendsFailWindow   { get; set; } = "3";
            public bool   TrendsAutoWatch    { get; set; } = false;
            public int    TrendsWatchIntervalSecs { get; set; } = 60;

            // ── Last-used folders (per feature) ──────────────────────────────
            public string LastNmonOutputDir  { get; set; } = "";
        }

        public static AppSettings LoadSettings()
        {
            try
            {
                if (!File.Exists(SettingsPath)) return new();
                var json = File.ReadAllText(SettingsPath);
                var s = JsonSerializer.Deserialize<AppSettings>(json, Opts);
                if (s == null) return new();
                return MigrateSettings(s, s.Version);
            }
            catch { return new(); }
        }

        public static void SaveSettings(AppSettings settings)
        {
            try
            {
                EnsureDir();
                settings.Version = CurrentSettingsVersion;
                File.WriteAllText(SettingsPath, JsonSerializer.Serialize(settings, Opts));
            }
            catch { }
        }

        private static AppSettings MigrateSettings(AppSettings s, int fromVersion)
        {
            // Add migration steps here as needed.
            return s;
        }

        // ─────────────────────────────────────────────────────────────────────
        // Global Backup — single file containing everything
        // ─────────────────────────────────────────────────────────────────────

        public class GlobalBackup
        {
            public int                          Version        { get; set; } = 1;
            public DateTime                     ExportedAt     { get; set; } = DateTime.Now;
            public ScriptLibraryFile            ScriptLibrary  { get; set; } = new();
            public TrendsLibraryFile            TrendsLibrary  { get; set; } = new();
            public AppSettings                  Settings       { get; set; } = new();
        }

        public static bool ExportAll(
            List<MainWindow.ScriptEntry> scripts,
            List<TrendsCustomerDto>      trends,
            AppSettings                  settings,
            string                       destPath)
        {
            try
            {
                var backup = new GlobalBackup
                {
                    ExportedAt    = DateTime.Now,
                    ScriptLibrary = new ScriptLibraryFile  { Version = CurrentScriptLibraryVersion, Entries = scripts },
                    TrendsLibrary = new TrendsLibraryFile  { Version = CurrentTrendsLibraryVersion, Entries = trends  },
                    Settings      = settings,
                };
                File.WriteAllText(destPath, JsonSerializer.Serialize(backup, Opts));
                return true;
            }
            catch { return false; }
        }

        public class GlobalImportResult
        {
            public List<MainWindow.ScriptEntry> NewScripts    { get; set; } = new();
            public List<TrendsCustomerDto>      NewCustomers  { get; set; } = new();
            public AppSettings?                 Settings      { get; set; }
            public int                          TotalScripts  { get; set; }
            public int                          TotalCustomers { get; set; }
            public string?                      Error         { get; set; }
        }

        /// <summary>
        /// Reads a global backup file and returns what is new vs what already exists.
        /// Callers should confirm with the user before applying.
        /// </summary>
        public static GlobalImportResult ImportAll(
            string                       sourcePath,
            List<MainWindow.ScriptEntry> existingScripts,
            List<TrendsCustomerDto>      existingCustomers)
        {
            var result = new GlobalImportResult();
            try
            {
                var json = File.ReadAllText(sourcePath);
                var backup = JsonSerializer.Deserialize<GlobalBackup>(json, Opts);
                if (backup == null) { result.Error = "File could not be parsed."; return result; }

                result.TotalScripts   = backup.ScriptLibrary.Entries.Count;
                result.TotalCustomers = backup.TrendsLibrary.Entries.Count;
                result.Settings       = backup.Settings;

                // Scripts — merge by Id
                var existingIds = new HashSet<string>(
                    System.Linq.Enumerable.Select(existingScripts, e => e.Id),
                    StringComparer.OrdinalIgnoreCase);
                result.NewScripts = backup.ScriptLibrary.Entries
                    .Where(e => !existingIds.Contains(e.Id))
                    .ToList();

                // Customers — merge by Name
                var existingNames = new HashSet<string>(
                    System.Linq.Enumerable.Select(existingCustomers, c => c.Name),
                    StringComparer.OrdinalIgnoreCase);
                result.NewCustomers = backup.TrendsLibrary.Entries
                    .Where(c => !existingNames.Contains(c.Name))
                    .ToList();
            }
            catch (Exception ex) { result.Error = ex.Message; }
            return result;
        }
    }
}
