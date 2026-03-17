using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace TestApp
{
    public enum BlgServerType { AppServer, DbServer }

    public class BlgConvertOptions
    {
        public string BlgPath { get; set; } = string.Empty;
        public BlgServerType ServerType { get; set; } = BlgServerType.AppServer;
        public string? CustomCounterFilePath { get; set; }
        public string? OutputDirectory { get; set; }
    }

    public static class BLGConverter
    {
        // ── Built-in counter templates (from supplied .txt files) ─────────────

        private static readonly string[] AppServerCounters =
        {
            @"\Memory\Available MBytes",
            @"\Memory\Pages Input/sec",
            @"\Memory\Pages Output/sec",
            @"\Network Interface(*)\Bytes Total/sec",
            @"\PhysicalDisk(_Total)\% Idle Time",
            @"\PhysicalDisk(_Total)\Avg. Disk sec/Transfer",
            @"\PhysicalDisk(_Total)\Current Disk Queue Length",
            @"\Paging File(_Total)\% Usage",
            @"\Process(*)\% Processor Time",
            @"\Process(*)\Working Set",
            @"\Processor(_Total)\% Processor Time",
            @"\PhysicalDisk(*)\*",
        };

        private static readonly string[] DbServerCounters =
        {
            @"\Memory\Available MBytes",
            @"\Memory\Pages Input/sec",
            @"\Memory\Pages Output/sec",
            @"\Network Interface(*)\Bytes Total/sec",
            @"\PhysicalDisk(_Total)\% Idle Time",
            @"\PhysicalDisk(_Total)\Avg. Disk sec/Transfer",
            @"\PhysicalDisk(_Total)\Current Disk Queue Length",
            @"\Paging File(_Total)\% Usage",
            @"\Process(*)\% Processor Time",
            @"\Process(*)\Working Set",
            @"\Processor(_Total)\% Processor Time",
            @"\PhysicalDisk(*)\*",
            @"\SQLServer:Buffer Manager\Buffer cache hit ratio",
            @"\SQLServer:Buffer Manager\Page life expectancy",
            @"\SQLServer:General Statistics\User Connections",
            @"\SQLServer:Latches\Average Latch Wait Time (ms)",
            @"\SQLServer:Locks(_Total)\Average Wait Time (ms)",
            @"\SQLServer:Locks(_Total)\Number of Deadlocks/sec",
        };

        // ── Public API ────────────────────────────────────────────────────────

        /// <summary>
        /// Runs relog.exe against the .blg file with the resolved counter filter.
        /// Returns the path of the generated CSV file.
        /// </summary>
        public static string ConvertToCsv(BlgConvertOptions opts)
        {
            if (!File.Exists(opts.BlgPath))
                throw new FileNotFoundException("BLG file not found.", opts.BlgPath);

            string relogPath = FindRelog();

            string outDir = string.IsNullOrWhiteSpace(opts.OutputDirectory)
                ? Path.GetDirectoryName(opts.BlgPath)!
                : opts.OutputDirectory;
            Directory.CreateDirectory(outDir);

            string csvPath = Path.Combine(
                outDir,
                Path.GetFileNameWithoutExtension(opts.BlgPath) + ".csv");

            bool usingTemp = true; // ResolveCounterFile always writes a sanitized temp file
            string counterFilePath = ResolveCounterFile(opts);

            try
            {
                RunRelog(relogPath, opts.BlgPath, counterFilePath, csvPath);
            }
            finally
            {
                if (usingTemp && File.Exists(counterFilePath))
                    File.Delete(counterFilePath);
            }

            if (!File.Exists(csvPath))
                throw new InvalidOperationException(
                    $"relog.exe completed but no CSV was produced at:\n{csvPath}");

            return csvPath;
        }

        /// <summary>Returns the list of counters that will be applied (for UI preview).</summary>
        public static IReadOnlyList<string> PreviewCounters(BlgConvertOptions opts)
        {
            if (!string.IsNullOrWhiteSpace(opts.CustomCounterFilePath)
                && File.Exists(opts.CustomCounterFilePath))
            {
                return File.ReadAllLines(opts.CustomCounterFilePath)
                           .Select(l => l.Trim())
                           .Where(l => !string.IsNullOrWhiteSpace(l))
                           .ToList();
            }
            return opts.ServerType == BlgServerType.DbServer
                ? DbServerCounters
                : AppServerCounters;
        }

        /// <summary>Builds the relog command string shown in the UI.</summary>
        public static string BuildCommandPreview(BlgConvertOptions opts)
        {
            string blgName = Path.GetFileName(opts.BlgPath);
            if (string.IsNullOrWhiteSpace(blgName)) blgName = "<no file selected>";
            string csvName = string.IsNullOrWhiteSpace(blgName)
                ? "<output>.csv"
                : Path.GetFileNameWithoutExtension(blgName) + ".csv";

            string cfSource = !string.IsNullOrWhiteSpace(opts.CustomCounterFilePath)
                ? Path.GetFileName(opts.CustomCounterFilePath)
                : opts.ServerType == BlgServerType.DbServer
                    ? "db_detailed_counters.txt"
                    : "detailed_counters.txt";

            return $"relog \"{blgName}\" -cf \"{cfSource}\" -f CSV -o \"{csvName}\"";
        }

        // ── Internal helpers ──────────────────────────────────────────────────

        private static string ResolveCounterFile(BlgConvertOptions opts)
        {
            IEnumerable<string> counters;

            if (!string.IsNullOrWhiteSpace(opts.CustomCounterFilePath))
            {
                if (!File.Exists(opts.CustomCounterFilePath))
                    throw new FileNotFoundException(
                        "Custom counter file not found.", opts.CustomCounterFilePath);

                // Read and sanitize — original files may contain trailing tabs/CR (\t\r)
                counters = File.ReadAllLines(opts.CustomCounterFilePath)
                               .Select(l => l.Trim())
                               .Where(l => !string.IsNullOrWhiteSpace(l));
            }
            else
            {
                counters = opts.ServerType == BlgServerType.DbServer
                    ? DbServerCounters
                    : AppServerCounters;
            }

            string tempPath = Path.Combine(
                Path.GetTempPath(), $"blg_cf_{Guid.NewGuid():N}.txt");
            // Use UTF-8 without BOM — relog.exe misreads the first counter line if a BOM is present,
            // causing \Memory\Available MBytes (first entry) to be silently skipped.
            File.WriteAllLines(tempPath, counters, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            return tempPath;
        }

        private static void RunRelog(
            string relogExe, string blgPath, string cfPath, string csvPath)
        {
            var args = $"\"{blgPath}\" -cf \"{cfPath}\" -f CSV -o \"{csvPath}\" -y";

            var psi = new ProcessStartInfo
            {
                FileName = relogExe,
                Arguments = args,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
            };

            using var proc = Process.Start(psi)
                ?? throw new InvalidOperationException("Failed to start relog.exe.");

            string stdout = proc.StandardOutput.ReadToEnd();
            string stderr = proc.StandardError.ReadToEnd();
            proc.WaitForExit();

            if (proc.ExitCode != 0)
                throw new InvalidOperationException(
                    $"relog.exe exited with code {proc.ExitCode}.\n\n" +
                    $"Output:\n{stdout}\n{stderr}".TrimEnd());
        }

        private static string FindRelog()
        {
            // 1. System32 (most common)
            string sys32 = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System), "relog.exe");
            if (File.Exists(sys32)) return sys32;

            // 2. SysWOW64
            string wow64 = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                "SysWOW64", "relog.exe");
            if (File.Exists(wow64)) return wow64;

            // 3. PATH
            foreach (var dir in (Environment.GetEnvironmentVariable("PATH") ?? "")
                                 .Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
            {
                var candidate = Path.Combine(dir.Trim(), "relog.exe");
                if (File.Exists(candidate)) return candidate;
            }

            throw new FileNotFoundException(
                "relog.exe was not found.\n\n" +
                "It ships with Windows and is normally at:\n" +
                @"  C:\Windows\System32\relog.exe" + "\n\n" +
                "Please ensure it is on your PATH and retry.");
        }
    }
}
