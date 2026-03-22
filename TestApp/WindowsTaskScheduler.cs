using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace TestApp
{
    /// <summary>
    /// Registers Windows Task Scheduler entries via schtasks.exe.
    /// Creates a small .bat launcher file to handle working dir, env vars,
    /// and log redirection reliably — avoids quoting issues in /TR.
    /// </summary>
    public static class WindowsTaskScheduler
    {
        private const string TaskFolder = @"\PerformanceTestUtilities\";

        private static readonly Dictionary<string, (string Runtime, string ArgsPrefix)> KnownTypes =
            new(StringComparer.OrdinalIgnoreCase)
            {
                { ".ps1", ("powershell.exe", "-ExecutionPolicy Bypass -File") },
                { ".py",  ("python",         "") },
                { ".jar", ("java",           "-jar") },
                { ".bat", ("cmd.exe",        "/c") },
                { ".cmd", ("cmd.exe",        "/c") },
                { ".js",  ("node",           "") },
                { ".sh",  ("bash",           "") },
            };

        // ── Public API ────────────────────────────────────────────────────────

        /// <summary>
        /// Creates a Windows Task Scheduler task.
        /// Also writes a launcher .bat that handles env vars, working dir and logging.
        /// Returns (success, taskName, errorMessage, logFilePath).
        /// </summary>
        public static (bool Ok, string TaskName, string Error, string LogFile) CreateTask(MainWindow.ScriptEntry entry)
        {
            if (entry.Schedule == null)
                return (false, "", "No schedule defined.", "");

            string safeName = Regex.Replace(entry.Name, @"[^\w ]", "_").Trim();
            string taskName = TaskFolder + safeName + "_" + entry.Id;

            // Resolve runtime
            string ext = Path.GetExtension(entry.ScriptPath).ToLowerInvariant();
            string runtime, argsPrefix;

            if (!string.IsNullOrEmpty(entry.Runtime))
            {
                var p = entry.Runtime.Split(new[] { ' ' }, 2);
                runtime    = p[0];
                argsPrefix = p.Length > 1 ? p[1] : "";
            }
            else if (KnownTypes.TryGetValue(ext, out var known))
            {
                runtime    = known.Runtime;
                argsPrefix = known.ArgsPrefix;
            }
            else
            {
                return (false, "", $"Unknown script type '{ext}'. Set a Runtime Override.", "");
            }

            // Paths
            string scriptDir = Path.GetDirectoryName(entry.ScriptPath) ?? "";
            string workDir   = string.IsNullOrEmpty(entry.WorkingDir) ? scriptDir : entry.WorkingDir;

            // Use custom log path if set, otherwise default next to script
            string logFile = !string.IsNullOrEmpty(entry.Schedule.LogFilePath)
                ? entry.Schedule.LogFilePath
                : Path.Combine(scriptDir,
                    Path.GetFileNameWithoutExtension(entry.ScriptPath) + "_scheduled.log");

            // Launcher .bat lives next to the script
            string launcherPath = Path.Combine(scriptDir,
                Path.GetFileNameWithoutExtension(entry.ScriptPath) + "_launcher.bat");

            // Write the launcher bat
            WriteLauncher(launcherPath, runtime, argsPrefix, entry.ScriptPath,
                entry.Arguments ?? "", workDir, logFile, entry.EnvVars);

            // Build schedule trigger
            string trigger = BuildTrigger(entry.Schedule, out string startDate);
            if (string.IsNullOrEmpty(trigger))
                return (false, "", "Unsupported schedule type.", "");

            // /TR just calls the launcher bat — no quoting nightmares
            var args = new List<string>
            {
                "/Create", "/F",
                "/TN", taskName,
                "/TR", $"\"{launcherPath}\"",
            };

            foreach (var t in trigger.Split(' ', StringSplitOptions.RemoveEmptyEntries))
                args.Add(t);

            args.Add("/ST");
            args.Add(entry.Schedule.TimeOfDay);

            if (!string.IsNullOrEmpty(startDate))
            {
                args.Add("/SD");
                args.Add(startDate);
            }

            var (exitCode, output, error) = Run(args.ToArray());

            return exitCode == 0
                ? (true, taskName, "", logFile)
                : (false, taskName, string.IsNullOrWhiteSpace(error) ? output : error, "");
        }

        public static void DeleteTask(string taskName)
        {
            try { Run(new[] { "/Delete", "/F", "/TN", taskName }); }
            catch { }
        }

        /// <summary>Deletes the launcher .bat file written alongside the script.</summary>
        public static void DeleteLauncher(string scriptPath)
        {
            try
            {
                string launcherPath = Path.Combine(
                    Path.GetDirectoryName(scriptPath) ?? "",
                    Path.GetFileNameWithoutExtension(scriptPath) + "_launcher.bat");
                if (File.Exists(launcherPath))
                    File.Delete(launcherPath);
            }
            catch { }
        }

        public static bool TaskExists(string taskName)
        {
            var (exit, _, _) = Run(new[] { "/Query", "/TN", taskName });
            return exit == 0;
        }

        // ── Launcher .bat writer ──────────────────────────────────────────────

        private static void WriteLauncher(
            string launcherPath,
            string runtime,
            string argsPrefix,
            string scriptPath,
            string userArgs,
            string workDir,
            string logFile,
            Dictionary<string, string>? envVars)
        {
            var bat = new System.Text.StringBuilder();
            bat.AppendLine("@echo off");
            bat.AppendLine("setlocal");
            bat.AppendLine();

            // Set env vars
            if (envVars != null)
                foreach (var kv in envVars)
                    bat.AppendLine($"set {kv.Key}={kv.Value}");

            // Set working directory
            if (!string.IsNullOrEmpty(workDir))
                bat.AppendLine($"cd /d \"{workDir}\"");

            bat.AppendLine();

            // Log header
            bat.AppendLine($"echo ======================================== >> \"{logFile}\"");
            bat.AppendLine($"echo Scheduled run: %DATE% %TIME% >> \"{logFile}\"");
            bat.AppendLine($"echo ======================================== >> \"{logFile}\"");
            bat.AppendLine();

            // Run the script — redirect stdout + stderr to log
            string call = string.IsNullOrEmpty(argsPrefix)
                ? $"\"{runtime}\" \"{scriptPath}\" {userArgs}".Trim()
                : $"\"{runtime}\" {argsPrefix} \"{scriptPath}\" {userArgs}".Trim();

            bat.AppendLine($"{call} >> \"{logFile}\" 2>&1");
            bat.AppendLine();
            bat.AppendLine($"echo Exit code: %ERRORLEVEL% >> \"{logFile}\"");
            bat.AppendLine("endlocal");

            File.WriteAllText(launcherPath, bat.ToString());
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static string BuildTrigger(MainWindow.ScriptSchedule sched, out string startDate)
        {
            startDate = "";
            switch (sched.Type)
            {
                case "Once":
                    if (!DateTime.TryParse(sched.RunOnce, out var runAt)) return "";
                    startDate = runAt.ToString(
                        System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern);
                    return "/SC ONCE";

                case "Daily":
                    return "/SC DAILY /MO 1";

                case "Weekly":
                    string day = ((DayOfWeek)sched.DayOfWeek).ToString().Substring(0, 3).ToUpper();
                    return "/SC WEEKLY /D " + day;

                default:
                    return "";
            }
        }

        private static (int ExitCode, string Output, string Error) Run(string[] args)
        {
            var psi = new ProcessStartInfo
            {
                FileName               = "schtasks.exe",
                UseShellExecute        = false,
                RedirectStandardOutput = true,
                RedirectStandardError  = true,
                CreateNoWindow         = true,
            };
            foreach (var a in args)
                psi.ArgumentList.Add(a);

            using var proc = Process.Start(psi)
                ?? throw new InvalidOperationException("Failed to start schtasks.exe");

            string output = proc.StandardOutput.ReadToEnd().Trim();
            string error  = proc.StandardError.ReadToEnd().Trim();
            proc.WaitForExit();
            return (proc.ExitCode, output, error);
        }

        // ── App auto-start on Windows logon ──────────────────────────────────
        //
        // Uses the HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run registry
        // key so the app restarts after a server reboot (or user logon) without
        // requiring admin rights.  The entry launches the app with --minimized
        // so it goes straight to the system tray and resumes watches silently.

        private const string AutoStartRegistryKey  = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        private const string AutoStartRegistryName = "PerformanceTestUtilities";

        /// <summary>
        /// Registers the app to auto-start on Windows logon via the
        /// HKCU Run registry key.  Returns (ok, error).
        /// </summary>
        public static (bool Ok, string Error) RegisterAppAutoStart()
        {
            try
            {
                string exePath = GetAppExePath();
                if (string.IsNullOrEmpty(exePath))
                    return (false, "Could not determine the application executable path.");

                using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(AutoStartRegistryKey, writable: true);
                if (key == null)
                    return (false, "Could not open HKCU Run registry key.");

                key.SetValue(AutoStartRegistryName, $"\"{exePath}\" --minimized");
                return (true, "");
            }
            catch (Exception ex) { return (false, ex.Message); }
        }

        /// <summary>Removes the app auto-start registry entry if it exists.</summary>
        public static (bool Ok, string Error) UnregisterAppAutoStart()
        {
            try
            {
                using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(AutoStartRegistryKey, writable: true);
                if (key == null) return (true, "");

                if (key.GetValue(AutoStartRegistryName) != null)
                    key.DeleteValue(AutoStartRegistryName, throwOnMissingValue: false);

                return (true, "");
            }
            catch (Exception ex) { return (false, ex.Message); }
        }

        /// <summary>Returns true if the app auto-start registry entry is present.</summary>
        public static bool IsAppAutoStartRegistered()
        {
            try
            {
                using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(AutoStartRegistryKey);
                return key?.GetValue(AutoStartRegistryName) != null;
            }
            catch { return false; }
        }

        /// <summary>
        /// Resolves the path to the running executable.  Works for both
        /// framework-dependent (.dll + dotnet host) and self-contained (.exe)
        /// deployments.
        /// </summary>
        private static string GetAppExePath()
        {
            // For single-file / self-contained: Environment.ProcessPath is the .exe
            string? processPath = Environment.ProcessPath;
            if (!string.IsNullOrEmpty(processPath) &&
                processPath.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                return processPath;

            // Fallback: look for TestApp.exe next to the entry assembly
            string? asmLocation = System.Reflection.Assembly.GetEntryAssembly()?.Location;
            if (!string.IsNullOrEmpty(asmLocation))
            {
                string dir = Path.GetDirectoryName(asmLocation) ?? "";
                string candidate = Path.Combine(dir, "TestApp.exe");
                if (File.Exists(candidate)) return candidate;
            }

            return processPath ?? "";
        }
    }
}
