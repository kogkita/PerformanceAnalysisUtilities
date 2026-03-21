using System;
using System.IO;

namespace TestApp
{
    /// <summary>
    /// Optional app-wide file logger.  Writes timestamped lines to a rolling
    /// daily log file inside the user-configured log folder.
    ///
    /// Enabled/disabled and the folder are controlled via AppSettings
    /// (AppLogEnabled + AppLogFolder).  All methods are thread-safe and
    /// never throw — failures are silently swallowed so the app keeps running.
    /// </summary>
    public static class AppLogger
    {
        private static readonly object _lock = new();
        private static bool   _enabled    = false;
        private static string _logFolder  = "";

        // ── Called once at startup and whenever settings are saved ────────────

        public static void Configure(bool enabled, string folder)
        {
            lock (_lock)
            {
                _enabled   = enabled;
                _logFolder = folder?.Trim() ?? "";
            }
        }

        // ── Write a single line ───────────────────────────────────────────────

        public static void Write(string source, string message)
        {
            bool   enabled;
            string folder;
            lock (_lock) { enabled = _enabled; folder = _logFolder; }

            if (!enabled || string.IsNullOrEmpty(folder)) return;

            try
            {
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);

                string fileName = $"AppLog_{DateTime.Now:yyyy-MM-dd}.txt";
                string path     = Path.Combine(folder, fileName);
                string line     = $"[{DateTime.Now:HH:mm:ss.fff}]  [{source}]  {message}";

                lock (_lock)
                    File.AppendAllText(path, line + Environment.NewLine);
            }
            catch { /* non-fatal */ }
        }

        // ── Convenience: write a block header ────────────────────────────────

        public static void WriteHeader(string source, string title)
        {
            Write(source, $"{'─',1}── {title} {'─',1}──────────────────────────────");
        }

        // ── Returns an Action<string> delegate bound to a source tag ─────────
        // Pass this into Generate() so each tool's log lines are tagged.

        public static Action<string> GetWriter(string source)
            => msg => Write(source, msg);
    }
}
