using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Chart;
using System.Drawing;
using System.Globalization;
using System.IO;

namespace TestApp
{
    // ── Data models ───────────────────────────────────────────────────────────

    public class TrendTestCase
    {
        public string Name { get; set; } = "";
        public string Status { get; set; } = "";   // PASS / FAIL / ""
        public int? Seconds { get; set; }         // null = not run this cycle
        public string TimeStr { get; set; } = "";   // original HH:mm:ss
    }

    public class TrendRun
    {
        public string RunNumber { get; set; } = "";
        public string Label { get; set; } = "";   // e.g. "TRN_FEB_26_2"
        public DateTime RunDate { get; set; }
        public List<TrendTestCase> Cases { get; set; } = new();

        // Computed
        public int Total => Cases.Count;
        public int Passed => Cases.Count(c => c.Status == "PASS");
        public int Failed => Cases.Count(c => c.Status == "FAIL");
        public double PassPct => Total > 0 ? Math.Round(Passed * 100.0 / Total, 1) : 0;
        public int TotalSeconds => Cases.Where(c => c.Seconds.HasValue).Sum(c => c.Seconds!.Value);
        public int AvgSeconds => Cases.Count(c => c.Seconds.HasValue) > 0
            ? TotalSeconds / Cases.Count(c => c.Seconds.HasValue) : 0;
    }

    public class TrendFlag
    {
        public string TestCase { get; set; } = "";
        public string Type { get; set; } = "";   // FAIL / NEW / MISSING / SLOWER / FASTER
        public string RunLabel { get; set; } = "";
        public string Detail { get; set; } = "";
    }

    // ── Main processor ────────────────────────────────────────────────────────

    public static class TestRunTrendsProcessor
    {
        public static Action<string>? Log { get; set; }

        private static void Write(string msg) => Log?.Invoke(msg);

        private static int ParseSecs(string? t)
        {
            if (string.IsNullOrEmpty(t)) return 0;
            var p = t.Split(':');
            if (p.Length == 3 &&
                int.TryParse(p[0], out int h) &&
                int.TryParse(p[1], out int m) &&
                int.TryParse(p[2], out int s))
                return h * 3600 + m * 60 + s;
            return 0;
        }

        private static string FmtTime(int secs)
        {
            if (secs <= 0) return "-";
            int h = secs / 3600, m = (secs % 3600) / 60, s = secs % 60;
            return h > 0 ? $"{h}h {m:D2}m" : $"{m}m {s:D2}s";
        }

        // ── Parse one input Excel file ────────────────────────────────────────

        public static TrendRun? ParseRunFile(string xlsxPath)
        {
            try
            {
                ExcelPackage.License.SetNonCommercialPersonal("Test Run Trends");

                // Copy to temp first — handles OneDrive/SharePoint locked files
                string tempPath = Path.Combine(Path.GetTempPath(),
                    "TrendRun_" + Path.GetFileName(xlsxPath));
                File.Copy(xlsxPath, tempPath, overwrite: true);

                using var pkg = new ExcelPackage(new FileInfo(tempPath));
                var ws = pkg.Workbook.Worksheets.FirstOrDefault();

                try { File.Delete(tempPath); } catch { }

                if (ws == null)
                {
                    Write($"  SKIP {Path.GetFileName(xlsxPath)}: no worksheets found");
                    return null;
                }

                // Find header row
                int headerRow = 1;
                string[]? headers = null;
                for (int r = 1; r <= Math.Min(5, ws.Dimension?.Rows ?? 1); r++)
                {
                    var h = Enumerable.Range(1, ws.Dimension?.Columns ?? 14)
                        .Select(c => ws.Cells[r, c].Text?.Trim() ?? "")
                        .ToArray();
                    if (h.Contains("Test Case Name") && h.Contains("Status"))
                    {
                        headerRow = r; headers = h; break;
                    }
                }
                if (headers == null) return null;

                int colCase = Array.IndexOf(headers, "Test Case Name") + 1;
                int colStatus = Array.IndexOf(headers, "Status") + 1;
                int colTime = Array.IndexOf(headers, "Total Time Taken") + 1;
                int colRun = Array.IndexOf(headers, "Run Number") + 1;
                int colDate = Array.IndexOf(headers, "Start Date") + 1;

                var run = new TrendRun();
                DateTime earliest = DateTime.MaxValue;

                int lastRow = ws.Dimension?.Rows ?? 1;
                for (int r = headerRow + 1; r <= lastRow; r++)
                {
                    string caseName = ws.Cells[r, colCase].Text?.Trim() ?? "";
                    if (string.IsNullOrEmpty(caseName)) continue;

                    string status = ws.Cells[r, colStatus].Text?.Trim() ?? "";
                    string timeStr = ws.Cells[r, colTime].Text?.Trim() ?? "";
                    string runNum = colRun > 0 ? ws.Cells[r, colRun].Text?.Trim() ?? "" : "";

                    if (string.IsNullOrEmpty(run.RunNumber) && !string.IsNullOrEmpty(runNum))
                        run.RunNumber = runNum;

                    if (colDate > 0)
                    {
                        var dateVal = ws.Cells[r, colDate].Value;
                        DateTime dt = DateTime.MinValue;
                        if (dateVal is DateTime d) dt = d;
                        else if (DateTime.TryParse(ws.Cells[r, colDate].Text, out dt)) { }
                        if (dt != DateTime.MinValue && dt < earliest) earliest = dt;
                    }

                    int secs = ParseSecs(timeStr);
                    run.Cases.Add(new TrendTestCase
                    {
                        Name = caseName,
                        Status = status,
                        Seconds = secs > 0 ? secs : null,
                        TimeStr = timeStr
                    });
                }

                run.RunDate = earliest != DateTime.MaxValue ? earliest : DateTime.MinValue;
                // Use friendly display label (e.g. "FEB 26") derived from run number
                string rawLabel = !string.IsNullOrEmpty(run.RunNumber) ? run.RunNumber
                    : Path.GetFileNameWithoutExtension(xlsxPath);
                run.Label = BuildDisplayLabel(rawLabel);
                if (run.Label == rawLabel) run.Label = rawLabel;  // fallback

                Write($"  Parsed: {run.Label} — {run.Total} cases, {run.Passed} pass, {run.Failed} fail, date {run.RunDate:dd MMM yyyy}");
                return run;
            }
            catch (Exception ex)
            {
                Write($"  ERROR parsing {Path.GetFileName(xlsxPath)}: {ex.Message}");
                return null;
            }
        }

        // ── Run label parsing helpers ─────────────────────────────────────────

        private static readonly Dictionary<string, int> MonthMap = new(StringComparer.OrdinalIgnoreCase)
        {
            {"JAN",1},{"FEB",2},{"MAR",3},{"APR",4},{"MAY",5},{"JUN",6},
            {"JUL",7},{"AUG",8},{"SEP",9},{"OCT",10},{"NOV",11},{"DEC",12}
        };

        /// <summary>Returns year*100+month key for grouping. 0 if unparseable.</summary>
        private static int ParseMonthYearKey(string label)
        {
            var parts = label.Split('_');
            for (int i = 0; i < parts.Length - 1; i++)
                if (MonthMap.TryGetValue(parts[i], out int m) &&
                    int.TryParse(parts[i + 1], out int y))
                {
                    int year = y < 100 ? 2000 + y : y;
                    return year * 100 + m;
                }
            return 0;
        }

        private static int ParseSortKey(string label) => ParseMonthYearKey(label);

        /// <summary>Returns the trailing run number (last numeric token), or 0.</summary>
        private static int ParseRunNumber(string label)
        {
            var parts = label.Split('_');
            for (int i = parts.Length - 1; i >= 0; i--)
                if (int.TryParse(parts[i], out int n)) return n;
            return 0;
        }

        /// <summary>Returns a friendly label like "FEB 26" from a run number string.</summary>
        private static string BuildDisplayLabel(string label)
        {
            var parts = label.Split('_');
            for (int i = 0; i < parts.Length - 1; i++)
                if (MonthMap.ContainsKey(parts[i].ToUpper()) &&
                    int.TryParse(parts[i + 1], out _))
                    return parts[i].ToUpper() + " " + parts[i + 1];
            return label;
        }

        // ── Main entry point ─────────────────────────────────────────────────

        public static (bool Ok, string OutputPath, string Error) Generate(
            string runsFolder, string customerName, string? reportsFolder = null, int failWindow = 3)
        {
            try
            {
                string outputFolder = string.IsNullOrEmpty(reportsFolder) ? runsFolder : reportsFolder;
                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                Write($"Scanning runs folder: {runsFolder}");

                // Find and parse all input Excel files (exclude trends output and manifest)
                string trendsFileName = customerName + "_Trends.xlsx";
                var inputFiles = Directory.GetFiles(runsFolder, "*.xlsx", SearchOption.TopDirectoryOnly)
                    .Where(f => !Path.GetFileName(f).Equals(trendsFileName, StringComparison.OrdinalIgnoreCase))
                    .OrderBy(f => f)
                    .ToList();

                if (inputFiles.Count == 0)
                    return (false, "", "No Excel files found in the selected folder.");

                Write($"Found {inputFiles.Count} file(s)...");
                var runs = new List<TrendRun>();
                foreach (var f in inputFiles)
                {
                    Write($"Reading {Path.GetFileName(f)}...");
                    var run = ParseRunFile(f);
                    if (run != null) runs.Add(run);
                }

                if (runs.Count == 0)
                    return (false, "", "No valid run files could be parsed.");

                // Deduplicate: for same Month+Year keep only highest RunNumber
                // Use RunNumber (original) not Label (display) for parsing
                runs = runs
                    .GroupBy(r => ParseMonthYearKey(r.RunNumber))
                    .Select(g => g.OrderByDescending(r => ParseRunNumber(r.RunNumber)).First())
                    .OrderByDescending(r => ParseSortKey(r.RunNumber))  // newest first (left) → oldest last (right)
                    .ToList();

                Write($"After deduplication: {runs.Count} unique month(s)");
                Write($"Loaded {runs.Count} run(s), building trends...");

                // Build trend output
                string outputPath = Path.Combine(outputFolder, trendsFileName);
                ExcelPackage.License.SetNonCommercialPersonal("Test Run Trends");
                using var pkg = new ExcelPackage();

                WriteExecutiveSummarySheet(pkg, runs, customerName);
                WriteTrendSheet(pkg, runs, failWindow);
                WriteFlagsSheet(pkg, runs);
                WriteChartsSheet(pkg, runs, customerName);

                pkg.SaveAs(new FileInfo(outputPath));
                Write($"Done — saved to: {outputPath}");
                return (true, outputPath, "");
            }
            catch (Exception ex)
            {
                return (false, "", ex.Message);
            }
        }

        // ── Sheet 1: Executive Summary ────────────────────────────────────────

        private static void WriteExecutiveSummarySheet(ExcelPackage pkg, List<TrendRun> runs, string customerName)
        {
            var ws = pkg.Workbook.Worksheets.Add("Summary");

            // Title
            ws.Cells["A1"].Value = $"{customerName} — Test Run Trend Summary";
            ws.Cells["A1"].Style.Font.Size = 16;
            ws.Cells["A1"].Style.Font.Bold = true;
            ws.Cells["A1"].Style.Font.Color.SetColor(Color.White);
            ws.Cells[1, 1, 1, 8].Merge = true;
            ws.Cells[1, 1, 1, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[1, 1, 1, 8].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x1E, 0x40, 0xAF));

            ws.Cells["A2"].Value = $"Generated: {DateTime.Now:dd MMM yyyy HH:mm}";
            ws.Cells["A2"].Style.Font.Color.SetColor(Color.FromArgb(0x6B, 0x7A, 0x99));
            ws.Cells["A2"].Style.Font.Italic = true;

            // Header row
            int hRow = 4;
            var headers = new[] { "Run", "Date", "Total Cases", "Passed", "Failed", "Pass %", "Total Runtime", "Avg per Test Case" };
            for (int c = 0; c < headers.Length; c++)
            {
                ws.Cells[hRow, c + 1].Value = headers[c];
                ws.Cells[hRow, c + 1].Style.Font.Bold = true;
                ws.Cells[hRow, c + 1].Style.Font.Color.SetColor(Color.White);
                ws.Cells[hRow, c + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[hRow, c + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x1E, 0x40, 0xAF));
                ws.Cells[hRow, c + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            // Data rows
            for (int i = 0; i < runs.Count; i++)
            {
                var r = runs[i];
                int row = hRow + 1 + i;
                bool alt = i % 2 == 1;
                var bg = alt ? Color.FromArgb(0xF0, 0xF4, 0xFF) : Color.White;

                void Cell(int col, object val, bool bold = false, Color? fg = null, ExcelHorizontalAlignment align = ExcelHorizontalAlignment.Left)
                {
                    ws.Cells[row, col].Value = val;
                    ws.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[row, col].Style.Fill.BackgroundColor.SetColor(bg);
                    if (bold) ws.Cells[row, col].Style.Font.Bold = true;
                    if (fg.HasValue) ws.Cells[row, col].Style.Font.Color.SetColor(fg.Value);
                    ws.Cells[row, col].Style.HorizontalAlignment = align;
                }

                Cell(1, r.Label, bold: true);
                Cell(2, r.RunDate == DateTime.MinValue ? "-" : r.RunDate.ToString("dd MMM yyyy"),
                    align: ExcelHorizontalAlignment.Center);
                Cell(3, r.Total, align: ExcelHorizontalAlignment.Center);
                Cell(4, r.Passed, fg: Color.FromArgb(0x16, 0x65, 0x34), align: ExcelHorizontalAlignment.Center);

                if (r.Failed > 0)
                {
                    ws.Cells[row, 5].Value = r.Failed;
                    ws.Cells[row, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xFF, 0xE4, 0xE4));
                    ws.Cells[row, 5].Style.Font.Color.SetColor(Color.FromArgb(0x99, 0x1B, 0x1B));
                    ws.Cells[row, 5].Style.Font.Bold = true;
                    ws.Cells[row, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                else Cell(5, 0, align: ExcelHorizontalAlignment.Center);

                Color passPctColor = r.PassPct >= 99 ? Color.FromArgb(0x16, 0x65, 0x34) :
                                     r.PassPct >= 95 ? Color.FromArgb(0x92, 0x40, 0x0E) :
                                     Color.FromArgb(0x99, 0x1B, 0x1B);
                Cell(6, $"{r.PassPct}%", fg: passPctColor, bold: true, align: ExcelHorizontalAlignment.Center);
                Cell(7, FmtTime(r.TotalSeconds), align: ExcelHorizontalAlignment.Center);
                Cell(8, FmtTime(r.AvgSeconds), align: ExcelHorizontalAlignment.Center);
            }

            // Border entire table
            int lastRow = hRow + runs.Count;
            var tableRange = ws.Cells[hRow, 1, lastRow, 8];
            tableRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            tableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            tableRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            tableRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;

            // Autofit
            for (int c = 1; c <= 8; c++) ws.Column(c).AutoFit();
            ws.Column(1).Width = Math.Max(ws.Column(1).Width, 28);
        }

        // ── Sheet 2: Per-test trend matrix ────────────────────────────────────

        private static void WriteTrendSheet(ExcelPackage pkg, List<TrendRun> runs, int failWindow = 3)
        {
            var ws = pkg.Workbook.Worksheets.Add("Test Case Trends");

            // Build lookup: runIndex -> caseName -> TrendTestCase
            var lookup = runs.Select(r =>
                r.Cases.ToDictionary(c => c.Name, c => c, StringComparer.OrdinalIgnoreCase)
            ).ToList();

            // Latest N runs for fail window
            int windowStart = Math.Max(0, runs.Count - failWindow);
            var windowRuns = lookup.Skip(windowStart).ToList();

            // Collect all unique test case names — sorted by fail count desc, then A-Z
            var allCases = runs.SelectMany(r => r.Cases.Select(c => c.Name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderByDescending(n => windowRuns.Count(wr => wr.TryGetValue(n, out var wc) && wc.Status == "FAIL"))
                .ThenBy(n => n)
                .ToList();

            // Single header row: Col A = Test Case Name, Col B = Fail count label,
            // Col C+ = one column per run/metric with "RunLabel\nStatus" / "RunLabel\nRuntime"
            void StyleHdr(ExcelRange cell, Color bg)
            {
                cell.Style.Font.Bold = true;
                cell.Style.Font.Color.SetColor(Color.White);
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(bg);
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell.Style.WrapText = true;
            }

            ws.Cells[1, 1].Value = "Test Case Name";
            StyleHdr(ws.Cells[1, 1], Color.FromArgb(0x1E, 0x40, 0xAF));

            ws.Cells[1, 2].Value = $"Failures in Last {Math.Min(failWindow, runs.Count)} Runs";
            StyleHdr(ws.Cells[1, 2], Color.FromArgb(0x1E, 0x40, 0xAF));

            for (int i = 0; i < runs.Count; i++)
            {
                int sc = 3 + i * 2;
                ws.Cells[1, sc].Value = $"{runs[i].Label}\nStatus";
                StyleHdr(ws.Cells[1, sc], Color.FromArgb(0x1E, 0x40, 0xAF));

                ws.Cells[1, sc + 1].Value = $"{runs[i].Label}\nRuntime";
                StyleHdr(ws.Cells[1, sc + 1], Color.FromArgb(0x1E, 0x40, 0xAF));
            }

            // Row height for the header to accommodate the two-line text
            ws.Row(1).Height = 30;

            // Data rows (now start at row 2, not row 3)
            for (int ri = 0; ri < allCases.Count; ri++)
            {
                string caseName = allCases[ri];
                int row = ri + 2;
                bool alt = ri % 2 == 1;

                ws.Cells[row, 1].Value = caseName;
                ws.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(
                    alt ? Color.FromArgb(0xF5, 0xF7, 0xFF) : Color.White);

                // Col B: fail count within the window
                int failsInWindow = windowRuns.Count(wr =>
                    wr.TryGetValue(caseName, out var wc) && wc.Status == "FAIL");
                ws.Cells[row, 2].Value = failsInWindow;
                ws.Cells[row, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                if (failsInWindow > 0)
                {
                    ws.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xFE, 0xE2, 0xE2));
                    ws.Cells[row, 2].Style.Font.Color.SetColor(Color.FromArgb(0x99, 0x1B, 0x1B));
                    ws.Cells[row, 2].Style.Font.Bold = true;
                }
                else
                {
                    ws.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(
                        alt ? Color.FromArgb(0xF5, 0xF7, 0xFF) : Color.White);
                }

                for (int i = 0; i < runs.Count; i++)
                {
                    int sc = 3 + i * 2;
                    var bg = alt ? Color.FromArgb(0xF5, 0xF7, 0xFF) : Color.White;

                    if (lookup[i].TryGetValue(caseName, out var tc))
                    {
                        // Status cell
                        ws.Cells[row, sc].Value = tc.Status;
                        ws.Cells[row, sc].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[row, sc].Style.Font.Bold = true;
                        ws.Cells[row, sc].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        if (tc.Status == "PASS")
                        {
                            ws.Cells[row, sc].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xD1, 0xFA, 0xE5));
                            ws.Cells[row, sc].Style.Font.Color.SetColor(Color.FromArgb(0x06, 0x5F, 0x46));
                        }
                        else if (tc.Status == "FAIL")
                        {
                            ws.Cells[row, sc].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xFE, 0xE2, 0xE2));
                            ws.Cells[row, sc].Style.Font.Color.SetColor(Color.FromArgb(0x99, 0x1B, 0x1B));
                        }
                        else
                        {
                            ws.Cells[row, sc].Style.Fill.BackgroundColor.SetColor(bg);
                        }

                        // Runtime cell — highlight if >25% slower than previous run
                        ws.Cells[row, sc + 1].Value = tc.TimeStr;
                        ws.Cells[row, sc + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[row, sc + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;

                        Color runtimeBg = bg;
                        if (i > 0 && tc.Seconds.HasValue && tc.Seconds > 0)
                        {
                            var prev = lookup[i - 1].TryGetValue(caseName, out var pt) ? pt : null;
                            if (prev?.Seconds > 0)
                            {
                                double pct = (tc.Seconds.Value - prev.Seconds.Value) * 100.0 / prev.Seconds.Value;
                                if (pct >= 25)
                                    runtimeBg = Color.FromArgb(0xFF, 0xF3, 0xCD);  // amber — slower
                                else if (pct <= -25)
                                    runtimeBg = Color.FromArgb(0xD1, 0xFA, 0xE5);  // green — faster
                            }
                        }
                        ws.Cells[row, sc + 1].Style.Fill.BackgroundColor.SetColor(runtimeBg);
                    }
                    else
                    {
                        // Not run this cycle
                        ws.Cells[row, sc].Value = "—";
                        ws.Cells[row, sc].Style.Font.Color.SetColor(Color.FromArgb(0xBB, 0xBB, 0xBB));
                        ws.Cells[row, sc].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[row, sc].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[row, sc].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xF3, 0xF4, 0xF6));
                        ws.Cells[row, sc + 1].Value = "—";
                        ws.Cells[row, sc + 1].Style.Font.Color.SetColor(Color.FromArgb(0xBB, 0xBB, 0xBB));
                        ws.Cells[row, sc + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[row, sc + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[row, sc + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xF3, 0xF4, 0xF6));
                    }
                }
            }

            ws.Column(1).Width = 52;
            ws.Column(2).Width = 14;   // Failures window col
            for (int i = 0; i < runs.Count; i++)
            {
                ws.Column(3 + i * 2).Width = 10;
                ws.Column(4 + i * 2).Width = 12;
            }

            // Autofilter on the single header row
            int lastCol = 2 + runs.Count * 2;
            string lastColLetter = GetColumnLetter(lastCol);
            ws.Cells[$"A1:{lastColLetter}1"].AutoFilter = true;

            ws.View.FreezePanes(2, 3);
        }

        // ── Sheet 3: Flags ────────────────────────────────────────────────────

        private static string GetColumnLetter(int col)
        {
            string letter = "";
            while (col > 0)
            {
                int mod = (col - 1) % 26;
                letter = (char)('A' + mod) + letter;
                col = (col - 1) / 26;
            }
            return letter;
        }

        private static void WriteFlagsSheet(ExcelPackage pkg, List<TrendRun> runs)
        {
            var ws = pkg.Workbook.Worksheets.Add("Flags");

            var flags = new List<TrendFlag>();

            var lookups = runs.Select(r =>
                r.Cases.ToDictionary(c => c.Name, c => c, StringComparer.OrdinalIgnoreCase)
            ).ToList();

            var allNames = runs.SelectMany(r => r.Cases.Select(c => c.Name))
                .Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            foreach (var name in allNames)
            {
                for (int i = 0; i < runs.Count; i++)
                {
                    lookups[i].TryGetValue(name, out var cur);
                    TrendTestCase? prev = i > 0 ? (lookups[i - 1].TryGetValue(name, out var p) ? p : null) : null;

                    if (cur == null && i > 0 && prev != null)
                        flags.Add(new TrendFlag { TestCase = name, Type = "MISSING", RunLabel = runs[i].Label, Detail = $"Was in {runs[i - 1].Label}, not in {runs[i].Label}" });

                    if (cur != null && i > 0 && prev == null)
                        flags.Add(new TrendFlag { TestCase = name, Type = "NEW", RunLabel = runs[i].Label, Detail = $"First appeared in {runs[i].Label}" });

                    if (cur?.Status == "FAIL")
                        flags.Add(new TrendFlag { TestCase = name, Type = "FAIL", RunLabel = runs[i].Label, Detail = $"Failed in {runs[i].Label}" });

                    if (cur?.Seconds > 0 && prev?.Seconds > 0)
                    {
                        double pct = (cur.Seconds!.Value - prev.Seconds!.Value) * 100.0 / prev.Seconds.Value;
                        if (pct >= 25)
                            flags.Add(new TrendFlag { TestCase = name, Type = "SLOWER", RunLabel = runs[i].Label, Detail = $"+{pct:0}% vs {runs[i - 1].Label} ({FmtTime(prev.Seconds.Value)} → {FmtTime(cur.Seconds.Value)})" });
                        else if (pct <= -25)
                            flags.Add(new TrendFlag { TestCase = name, Type = "FASTER", RunLabel = runs[i].Label, Detail = $"{pct:0}% vs {runs[i - 1].Label} ({FmtTime(prev.Seconds.Value)} → {FmtTime(cur.Seconds.Value)})" });
                    }
                }
            }

            // Sort: FAIL first, then MISSING, NEW, SLOWER, FASTER
            var order = new Dictionary<string, int> { { "FAIL", 0 }, { "MISSING", 1 }, { "NEW", 2 }, { "SLOWER", 3 }, { "FASTER", 4 } };
            flags = flags.OrderBy(f => order.GetValueOrDefault(f.Type, 9)).ThenBy(f => f.TestCase).ToList();

            // Header
            var headers = new[] { "Type", "Test Case", "Run", "Detail" };
            var hColors = new Dictionary<string, Color>
            {
                {"FAIL",    Color.FromArgb(0xFF, 0xE4, 0xE4)},
                {"MISSING", Color.FromArgb(0xFF, 0xF3, 0xCD)},
                {"NEW",     Color.FromArgb(0xD1, 0xFA, 0xE5)},
                {"SLOWER",  Color.FromArgb(0xFF, 0xF3, 0xCD)},
                {"FASTER",  Color.FromArgb(0xD1, 0xFA, 0xE5)},
            };

            for (int c = 0; c < headers.Length; c++)
            {
                ws.Cells[1, c + 1].Value = headers[c];
                ws.Cells[1, c + 1].Style.Font.Bold = true;
                ws.Cells[1, c + 1].Style.Font.Color.SetColor(Color.White);
                ws.Cells[1, c + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[1, c + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x1E, 0x40, 0xAF));
            }

            for (int i = 0; i < flags.Count; i++)
            {
                var f = flags[i];
                int row = i + 2;
                var bg = hColors.GetValueOrDefault(f.Type, Color.White);

                ws.Cells[row, 1].Value = f.Type;
                ws.Cells[row, 1].Style.Font.Bold = true;
                ws.Cells[row, 2].Value = f.TestCase;
                ws.Cells[row, 3].Value = f.RunLabel;
                ws.Cells[row, 4].Value = f.Detail;

                for (int c = 1; c <= 4; c++)
                {
                    ws.Cells[row, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[row, c].Style.Fill.BackgroundColor.SetColor(bg);
                }
            }

            ws.Column(1).Width = 12;
            ws.Column(2).Width = 52;
            ws.Column(3).Width = 22;
            ws.Column(4).Width = 55;
        }

        // ── Sheet 4: Charts ───────────────────────────────────────────────────

        private static void WriteChartsSheet(ExcelPackage pkg, List<TrendRun> runs, string customerName)
        {
            var ws = pkg.Workbook.Worksheets.Add("Charts");

            // Direction label at top
            string dirNote = runs.Count > 0
                ? $"X-axis: newest ({runs[0].Label}) on the left → oldest ({runs[^1].Label}) on the right"
                : "X-axis: newest → oldest";
            ws.Cells[1, 1].Value = dirNote;
            ws.Cells[1, 1].Style.Font.Italic = true;
            ws.Cells[1, 1].Style.Font.Color.SetColor(Color.FromArgb(0x6B, 0x7A, 0x99));
            ws.Cells[1, 1].Style.Font.Size = 10;
            ws.Cells[1, 1, 1, 16].Merge = true;

            // Chart data (row 60+ to stay clear of charts)
            int dataStartRow = 60;
            ws.Cells[dataStartRow, 1].Value = "Run";
            ws.Cells[dataStartRow, 2].Value = "Pass %";
            ws.Cells[dataStartRow, 3].Value = "Total Runtime (hrs)";
            ws.Cells[dataStartRow, 4].Value = "Failed";
            ws.Cells[dataStartRow, 5].Value = "Label";

            for (int i = 0; i < runs.Count; i++)
            {
                int r = dataStartRow + 1 + i;
                double hrs = Math.Round(runs[i].TotalSeconds / 3600.0, 1);
                ws.Cells[r, 1].Value = runs[i].Label;
                ws.Cells[r, 2].Value = runs[i].PassPct;
                ws.Cells[r, 3].Value = hrs;
                ws.Cells[r, 4].Value = runs[i].Failed;
                ws.Cells[r, 5].Value = $"{runs[i].Label}  |  {hrs}h / {runs[i].Total} TC";
            }

            int dataEndRow = dataStartRow + runs.Count;

            // ── Chart 1: Pass % trend (row 2, col A) ──────────────────────────
            var chart1 = ws.Drawings.AddChart("PassPctTrend", eChartType.LineMarkers) as ExcelLineChart;
            if (chart1 != null)
            {
                chart1.Title.Text = "Pass % Trend  (newest → oldest)";
                chart1.SetPosition(1, 5, 0, 5);
                chart1.SetSize(600, 300);
                var s = chart1.Series.Add(
                    ws.Cells[dataStartRow + 1, 2, dataEndRow, 2],
                    ws.Cells[dataStartRow + 1, 5, dataEndRow, 5]);
                s.Header = "Pass %";
                chart1.DataLabel.ShowValue = true;
                chart1.DataLabel.Position = eLabelPosition.Top;
                chart1.YAxis.MajorGridlines.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
                chart1.YAxis.MajorGridlines.Fill.Color = Color.FromArgb(0xE5, 0xE7, 0xEB);
                chart1.YAxis.MajorGridlines.Width = 0.5;
            }

            // ── Chart 2: Total Runtime trend (row 2, col J) ───────────────────
            var chart2 = ws.Drawings.AddChart("RuntimeTrend", eChartType.LineMarkers) as ExcelLineChart;
            if (chart2 != null)
            {
                chart2.Title.Text = "Total Runtime Trend  (newest → oldest)";
                chart2.SetPosition(1, 5, 9, 5);
                chart2.SetSize(600, 300);
                var s = chart2.Series.Add(
                    ws.Cells[dataStartRow + 1, 3, dataEndRow, 3],
                    ws.Cells[dataStartRow + 1, 5, dataEndRow, 5]);
                s.Header = "Runtime (hrs)";
                chart2.DataLabel.ShowValue = true;
                chart2.DataLabel.Position = eLabelPosition.Top;
                chart2.YAxis.MajorGridlines.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
                chart2.YAxis.MajorGridlines.Fill.Color = Color.FromArgb(0xE5, 0xE7, 0xEB);
                chart2.YAxis.MajorGridlines.Width = 0.5;
            }

            // ── Chart 3: Fail count bar (row 20, col A) ───────────────────────
            var chart3 = ws.Drawings.AddChart("FailCount", eChartType.ColumnClustered) as ExcelBarChart;
            if (chart3 != null)
            {
                chart3.Title.Text = "Failed Test Cases per Run  (newest → oldest)";
                chart3.SetPosition(20, 5, 0, 5);
                chart3.SetSize(600, 300);
                var s = chart3.Series.Add(
                    ws.Cells[dataStartRow + 1, 4, dataEndRow, 4],
                    ws.Cells[dataStartRow + 1, 1, dataEndRow, 1]);
                s.Header = "Failed";
                chart3.DataLabel.ShowValue = true;
                chart3.DataLabel.Position = eLabelPosition.OutEnd;
                chart3.YAxis.MajorGridlines.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
                chart3.YAxis.MajorGridlines.Fill.Color = Color.FromArgb(0xE5, 0xE7, 0xEB);
                chart3.YAxis.MajorGridlines.Width = 0.5;
            }


        }
    }
}