using System;
using System.Collections.Concurrent;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace TestApp
{
    /// <summary>
    /// Optional app-wide file logger.
    ///
    /// Design goals:
    ///   • Thread-safe — callable from any thread without blocking the caller.
    ///   • Buffered — messages are queued and flushed by a background thread,
    ///     so UI and worker threads are never stalled on file I/O.
    ///   • Rolling daily files — one file per day (AppLog_yyyy-MM-dd.txt).
    ///   • Max file size — when a daily file exceeds MaxFileSizeBytes it is
    ///     renamed with a _1 / _2 suffix and a fresh file is started.
    ///   • Auto-cleanup — log files older than RetentionDays are deleted on
    ///     startup and once per day thereafter.
    ///   • Never throws — all failures are silently swallowed.
    /// </summary>
    public static class AppLogger
    {
        // ── Configuration constants ───────────────────────────────────────────
        public const long MaxFileSizeBytes = 5 * 1024 * 1024;   // 5 MB per segment
        public const int  RetentionDays    = 30;                  // keep 30 days of logs
        public const int  FlushIntervalMs  = 2_000;               // flush queue every 2 s

        // ── Internal state ────────────────────────────────────────────────────
        private static readonly object   _cfgLock   = new();
        private static bool              _enabled   = false;
        private static string            _logFolder = "";

        private static readonly ConcurrentQueue<string> _queue = new();
        private static          CancellationTokenSource  _cts   = new();
        private static          Task?                    _flushTask;
        private static          DateTime                 _lastCleanup = DateTime.MinValue;

        // ── Public API ────────────────────────────────────────────────────────

        /// <summary>
        /// Called at startup and whenever settings change.
        /// Starts or stops the background flush task as needed.
        /// </summary>
        public static void Configure(bool enabled, string folder)
        {
            lock (_cfgLock)
            {
                _enabled   = enabled;
                _logFolder = folder?.Trim() ?? "";
            }

            if (enabled && !string.IsNullOrEmpty(folder))
                EnsureFlushTaskRunning();
            else
                StopFlushTask();
        }

        /// <summary>Enqueue a log message. Never blocks the caller.</summary>
        public static void Write(string source, string message)
        {
            bool enabled;
            lock (_cfgLock) { enabled = _enabled; }
            if (!enabled) return;

            _queue.Enqueue($"[{DateTime.Now:HH:mm:ss.fff}]  [{source}]  {message}");
        }

        /// <summary>Returns an Action&lt;string&gt; delegate bound to a source tag.</summary>
        public static Action<string> GetWriter(string source)
            => msg => Write(source, msg);

        /// <summary>
        /// Flush remaining queued messages and stop the background task.
        /// Call from Application.Exit / OnClosed.
        /// </summary>
        public static void Shutdown()
        {
            StopFlushTask();
            FlushQueue(); // drain anything left
        }

        // ── Background flush task ─────────────────────────────────────────────

        private static void EnsureFlushTaskRunning()
        {
            if (_flushTask != null && !_flushTask.IsCompleted) return;

            _cts = new CancellationTokenSource();
            var token = _cts.Token;

            _flushTask = Task.Run(async () =>
            {
                while (!token.IsCancellationRequested)
                {
                    try { await Task.Delay(FlushIntervalMs, token); } catch { break; }
                    FlushQueue();
                    MaybeRunCleanup();
                }
                FlushQueue(); // final drain on shutdown
            });
        }

        private static void StopFlushTask()
        {
            try { _cts.Cancel(); } catch { }
        }

        // ── Queue drain ───────────────────────────────────────────────────────

        private static void FlushQueue()
        {
            if (_queue.IsEmpty) return;

            string folder;
            lock (_cfgLock) { folder = _logFolder; }
            if (string.IsNullOrEmpty(folder)) return;

            try
            {
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);

                string baseName = $"AppLog_{DateTime.Now:yyyy-MM-dd}";
                string filePath = ResolveLogPath(folder, baseName);

                // Drain entire queue in one StreamWriter open/close
                using var sw = new StreamWriter(filePath, append: true);
                while (_queue.TryDequeue(out var line))
                    sw.WriteLine(line);
            }
            catch
            {
                // If we cannot write, drop the messages — never crash the app
                while (_queue.TryDequeue(out _)) { }
            }
        }

        /// <summary>
        /// Returns the active log file path for today, rolling to a new segment
        /// if the current file has exceeded MaxFileSizeBytes.
        /// </summary>
        private static string ResolveLogPath(string folder, string baseName)
        {
            string path = Path.Combine(folder, baseName + ".txt");
            if (!File.Exists(path)) return path;

            var fi = new FileInfo(path);
            if (fi.Length < MaxFileSizeBytes) return path;

            // Roll: rename current file with _N suffix, return original path (fresh)
            int seg = 1;
            string rolled;
            do { rolled = Path.Combine(folder, $"{baseName}_{seg++}.txt"); }
            while (File.Exists(rolled));

            try { File.Move(path, rolled); } catch { }
            return path;
        }

        // ── Old file cleanup ──────────────────────────────────────────────────

        private static void MaybeRunCleanup()
        {
            if ((DateTime.Now - _lastCleanup).TotalHours < 24) return;
            _lastCleanup = DateTime.Now;

            string folder;
            lock (_cfgLock) { folder = _logFolder; }
            if (string.IsNullOrEmpty(folder)) return;

            try
            {
                var cutoff = DateTime.Now.AddDays(-RetentionDays);
                foreach (var file in Directory.GetFiles(folder, "AppLog_*.txt"))
                {
                    try
                    {
                        if (File.GetLastWriteTime(file) < cutoff)
                            File.Delete(file);
                    }
                    catch { }
                }
            }
            catch { }
        }
    }
}
