using OfficeOpenXml;

using OfficeOpenXml.Style;

using System.Drawing;

using System.Globalization;

using System.IO;



namespace TestApp

{

    // ── Shared metric bag used for both CSV and JTL runs ─────────────────────



    public class ComparisonRecord

    {

        public string TransactionName { get; set; } = string.Empty;



        // Baseline

        public double BaseAvg { get; set; }

        public double BaseMedian { get; set; }

        public double BaseP90 { get; set; }

        public double BaseMin { get; set; }

        public double BaseMax { get; set; }

        public double BaseErrors { get; set; }   // 0-100

        public int BaseSamples { get; set; }



        // Current

        public double CurAvg { get; set; }

        public double CurMedian { get; set; }

        public double CurP90 { get; set; }

        public double CurMin { get; set; }

        public double CurMax { get; set; }

        public double CurErrors { get; set; }

        public int CurSamples { get; set; }



        // Delta helpers (positive = current is SLOWER / worse)

        public double DeltaAvgMs => CurAvg - BaseAvg;

        public double DeltaP90Ms => CurP90 - BaseP90;

        public double DeltaMedianMs => CurMedian - BaseMedian;

        public double DeltaErrorPct => CurErrors - BaseErrors;



        public double DeltaAvgPct => BaseAvg > 0 ? (CurAvg - BaseAvg) / BaseAvg * 100 : 0;

        public double DeltaP90Pct => BaseP90 > 0 ? (CurP90 - BaseP90) / BaseP90 * 100 : 0;



        public bool OnlyInBaseline { get; set; }

        public bool OnlyInCurrent { get; set; }

    }



    // ── Input type ────────────────────────────────────────────────────────────



    public enum ComparisonFileType { Csv, Jtl }



    // ── Comparison mode ───────────────────────────────────────────────────────

    // AllVsBaseline : Run2 vs Baseline, Run3 vs Baseline, Run4 vs Baseline …

    // Sequential    : Run2 vs Baseline, Run3 vs Run2, Run4 vs Run3 …



    public enum ComparisonMode { AllVsBaseline, Sequential }



    // ── Main processing class ─────────────────────────────────────────────────



    public static class RunComparisonProcessor

    {

        private const double RegressionThresholdPct = 10.0;



        // ── Colour palette — light, readable on white ─────────────────────────

        private static readonly Color ClrRegFill = Color.FromArgb(0xFF, 0xE4, 0xE4); // soft red

        private static readonly Color ClrRegText = Color.FromArgb(0x9B, 0x1C, 0x1C); // dark red

        private static readonly Color ClrImpFill = Color.FromArgb(0xDC, 0xFC, 0xE7); // soft green

        private static readonly Color ClrImpText = Color.FromArgb(0x16, 0x65, 0x34); // dark green

        private static readonly Color ClrStaFill = Color.FromArgb(0xF9, 0xFA, 0xFB); // near-white

        private static readonly Color ClrStaText = Color.FromArgb(0x37, 0x41, 0x51); // dark gray

        private static readonly Color ClrNewFill = Color.FromArgb(0xDB, 0xEA, 0xFE); // soft blue

        private static readonly Color ClrNewText = Color.FromArgb(0x1E, 0x3A, 0x8A); // dark blue

        private static readonly Color ClrRemFill = Color.FromArgb(0xFE, 0xF3, 0xC7); // soft amber

        private static readonly Color ClrRemText = Color.FromArgb(0x78, 0x35, 0x00); // dark amber

        private static readonly Color ClrHdrFill = Color.FromArgb(0x1E, 0x40, 0xAF); // blue header

        private static readonly Color ClrHdrText = Color.White;

        private static readonly Color ClrAltRow = Color.FromArgb(0xF3, 0xF4, 0xF6); // alternating row



        // ── Entry point ───────────────────────────────────────────────────────



        // ── Entry point: accepts 2..N run files ───────────────────────────────

        /// <summary>
        /// Auto-detect variant — determines CSV vs JTL per file from extension.
        /// </summary>
        public static void Compare(
            IList<string> runPaths,
            string outputPath,
            double slaThresholdMs = 0,
            ComparisonMode mode = ComparisonMode.AllVsBaseline)
        {
            if (runPaths == null || runPaths.Count < 2)
                throw new ArgumentException("At least two run files are required.");

            ExcelPackage.License.SetNonCommercialPersonal("Run Comparison");

            // Load all files — auto-detect type per file
            var allRecords = runPaths.Select(p => LoadRecords(p)).ToList();

            CompareInternal(runPaths, allRecords, outputPath, slaThresholdMs, mode);
        }

        /// <summary>
        /// Explicit file-type variant — all files are treated as the given type.
        /// Kept for backward compatibility.
        /// </summary>

        public static void Compare(

            IList<string> runPaths,

            string outputPath,

            ComparisonFileType fileType,

            double slaThresholdMs = 0,

            ComparisonMode mode = ComparisonMode.AllVsBaseline)

        {

            if (runPaths == null || runPaths.Count < 2)

                throw new ArgumentException("At least two run files are required.");



            ExcelPackage.License.SetNonCommercialPersonal("Run Comparison");



            // Load all files

            var allRecords = runPaths.Select(p => LoadRecords(p, fileType)).ToList();

            CompareInternal(runPaths, allRecords, outputPath, slaThresholdMs, mode);
        }

        private static void CompareInternal(
            IList<string> runPaths,
            List<List<FlatRecord>> allRecords,
            string outputPath,
            double slaThresholdMs,
            ComparisonMode mode)
        {
            // Build run labels: index 0 = "Run 1 (Baseline)", 1 = "Run 2", 2 = "Run 3" …
            var runLabels = new List<string>();
            for (int i = 0; i < runPaths.Count; i++)
                runLabels.Add(i == 0 ? "Run 1 (Baseline)" : $"Run {i + 1}");

            // Each entry: (leftLabel, rightLabel, leftRecords, rightRecords, delta, sheetPrefix)
            var pairs = new List<(string LeftLabel, string RightLabel,
                                  List<FlatRecord> LeftRec, List<FlatRecord> RightRec,
                                  List<ComparisonRecord> Delta, string SheetPrefix)>();

            if (mode == ComparisonMode.AllVsBaseline)
            {
                // Run 1 (Baseline) vs Run 2, Run 1 (Baseline) vs Run 3 …
                for (int i = 1; i < runPaths.Count; i++)
                {
                    var delta = BuildComparison(allRecords[0], allRecords[i]);
                    pairs.Add((runLabels[0], runLabels[i], allRecords[0], allRecords[i], delta, $"Run {i + 1} vs Run 1 "));
                }
            }
            else // Sequential
            {
                // Run 1 vs Run 2, Run 2 vs Run 3, Run 3 vs Run 4 …
                for (int i = 1; i < runPaths.Count; i++)
                {
                    var delta = BuildComparison(allRecords[i - 1], allRecords[i]);
                    pairs.Add((runLabels[i - 1], runLabels[i], allRecords[i - 1], allRecords[i], delta, $"Run {i + 1} vs Run {i} "));
                }
            }

            using var package = new ExcelPackage();

            // Summary — one section per pair
            WriteSummarySheet(package, pairs, runLabels, runPaths, mode, slaThresholdMs);

            // Avg + P90 sheets per pair
            foreach (var p in pairs)
            {
                WriteAvgSheet(package, p.Delta, slaThresholdMs, p.SheetPrefix, p.LeftLabel, p.RightLabel);
                WriteP90Sheet(package, p.Delta, slaThresholdMs, p.SheetPrefix, p.LeftLabel, p.RightLabel);
            }



            // Raw sheets — one per unique run file

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < runPaths.Count; i++)

            {

                if (seen.Add(runPaths[i]))

                {

                    string label = i == 0 ? "Baseline" : $"Run {i + 1}";

                    WriteRawSheet(package, allRecords[i], label);

                }

            }



            package.SaveAs(new FileInfo(outputPath));

        }



        // ── Legacy 2-file overload ────────────────────────────────────────────

        public static void Compare(

            string baselinePath,

            string currentPath,

            string outputPath,

            ComparisonFileType fileType,

            double slaThresholdMs = 0)

            => Compare(new[] { baselinePath, currentPath }, outputPath, fileType, slaThresholdMs);



        // ── Record loading ────────────────────────────────────────────────────



        /// <summary>
        /// Auto-detects file type from extension and loads records.
        /// .jtl → JTL parser, everything else → CSV parser.
        /// </summary>
        private static List<FlatRecord> LoadRecords(string path)
        {
            if (path.EndsWith(".jtl", StringComparison.OrdinalIgnoreCase))
                return LoadRecords(path, ComparisonFileType.Jtl);
            return LoadRecords(path, ComparisonFileType.Csv);
        }


        private static List<FlatRecord> LoadRecords(string path, ComparisonFileType fileType)

        {

            if (fileType == ComparisonFileType.Jtl)

            {

                return JTLFileProcessing.ParseJtl(path)

                    .Select(r => new FlatRecord

                    {

                        Name = r.TransactionName,

                        Average = r.Average,

                        Median = r.Median,

                        P90 = r.P90,

                        Min = r.Min,

                        Max = r.Max,

                        Errors = r.ErrorPercent,

                        Samples = r.Samples

                    }).ToList();

            }

            return LoadCsvRecords(path);

        }



        private static List<FlatRecord> LoadCsvRecords(string csvPath)

        {

            if (!File.Exists(csvPath))

                throw new FileNotFoundException("CSV file not found", csvPath);



            var lines = File.ReadAllLines(csvPath);

            if (lines.Length == 0)

                throw new InvalidDataException("CSV file is empty.");

            var headers = SplitCsvLine(lines[0]);



            int labelIdx = Array.IndexOf(headers, "Label");

            int sampIdx = Array.IndexOf(headers, "# Samples");

            int avgIdx = Array.IndexOf(headers, "Average");

            int medIdx = Array.IndexOf(headers, "Median");

            int minIdx = Array.IndexOf(headers, "Min");

            int maxIdx = Array.IndexOf(headers, "Max");

            int errIdx = Array.IndexOf(headers, "Error %");

            int p90Idx = -1;

            for (int i = 0; i < headers.Length; i++)

                if (headers[i].Contains("90%") || headers[i].Contains("90 %"))

                { p90Idx = i; break; }



            if (labelIdx < 0 || sampIdx < 0 || avgIdx < 0 ||

                medIdx < 0 || minIdx < 0 || maxIdx < 0 || errIdx < 0)

                throw new InvalidDataException(

                    "CSV file is missing one or more required columns (Label, # Samples, Average, Median, Min, Max, Error %).");



            var records = new List<FlatRecord>();

            for (int i = 1; i < lines.Length; i++)

            {

                if (string.IsNullOrWhiteSpace(lines[i])) continue;

                var v = SplitCsvLine(lines[i]);

                if (v.Length <= labelIdx) continue;

                string label = v[labelIdx].Trim();

                if (label.Equals("TOTAL", StringComparison.OrdinalIgnoreCase)) continue;

                if (label.StartsWith("/") || label.StartsWith("http")) continue;



                records.Add(new FlatRecord

                {

                    Name = label,

                    Samples = sampIdx < v.Length ? ParseInt(v[sampIdx]) : 0,

                    Average = avgIdx < v.Length ? ParseMs(v[avgIdx]) : 0,

                    Median = medIdx < v.Length ? ParseMs(v[medIdx]) : 0,

                    P90 = p90Idx >= 0 && p90Idx < v.Length ? ParseMs(v[p90Idx]) : 0,

                    Min = minIdx < v.Length ? ParseMs(v[minIdx]) : 0,

                    Max = maxIdx < v.Length ? ParseMs(v[maxIdx]) : 0,

                    Errors = errIdx < v.Length ? ParsePct(v[errIdx]) : 0

                });

            }

            return records;

        }



        // ── Delta computation ─────────────────────────────────────────────────



        private static List<ComparisonRecord> BuildComparison(

            List<FlatRecord> baseline, List<FlatRecord> current)

        {

            // Use GroupBy + First to gracefully handle duplicate transaction names
            // (ToDictionary would throw an ArgumentException on duplicates)
            var baseDict = baseline
                .GroupBy(r => r.Name, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            var curDict = current
                .GroupBy(r => r.Name, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);



            return baseDict.Keys

                .Union(curDict.Keys, StringComparer.OrdinalIgnoreCase)

                .OrderBy(n => n, StringComparer.Ordinal)

                .Select(name =>

                {

                    bool inBase = baseDict.TryGetValue(name, out var b);

                    bool inCur = curDict.TryGetValue(name, out var c);

                    return new ComparisonRecord

                    {

                        TransactionName = name,

                        OnlyInBaseline = inBase && !inCur,

                        OnlyInCurrent = !inBase && inCur,

                        BaseAvg = inBase ? b!.Average : 0,

                        BaseMedian = inBase ? b!.Median : 0,

                        BaseP90 = inBase ? b!.P90 : 0,

                        BaseMin = inBase ? b!.Min : 0,

                        BaseMax = inBase ? b!.Max : 0,

                        BaseErrors = inBase ? b!.Errors : 0,

                        BaseSamples = inBase ? b!.Samples : 0,

                        CurAvg = inCur ? c!.Average : 0,

                        CurMedian = inCur ? c!.Median : 0,

                        CurP90 = inCur ? c!.P90 : 0,

                        CurMin = inCur ? c!.Min : 0,

                        CurMax = inCur ? c!.Max : 0,

                        CurErrors = inCur ? c!.Errors : 0,

                        CurSamples = inCur ? c!.Samples : 0,

                    };

                })

                .ToList();

        }



        // ── Summary sheet ─────────────────────────────────────────────────────



        // ── Summary sheet ─────────────────────────────────────────────────────

        private static void WriteSummarySheet(
            ExcelPackage pkg,
            IList<(string LeftLabel, string RightLabel, List<FlatRecord> LeftRec, List<FlatRecord> RightRec, List<ComparisonRecord> Delta, string SheetPrefix)> pairs,
            IList<string> runLabels,
            IList<string> runPaths,
            ComparisonMode mode,
            double slaMs)
        {
            var ws = pkg.Workbook.Worksheets.Add("Summary");

            ws.Cells[1, 1].Value = "Run Comparison Report";
            ws.Cells[1, 1].Style.Font.Bold = true;
            ws.Cells[1, 1].Style.Font.Size = 16;
            ws.Cells[1, 1].Style.Font.Color.SetColor(ClrHdrFill);

            // Show "Run 1 (Baseline)   filename.csv" etc.
            for (int i = 0; i < runLabels.Count; i++)
                SetMeta(ws, 2 + i, runLabels[i], Path.GetFileName(runPaths[i]));

            int genRow = runLabels.Count + 2;
            SetMeta(ws, genRow, "Generated", DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
            SetMeta(ws, genRow + 1, "Mode", mode == ComparisonMode.Sequential
                ? "Sequential — each run vs previous (Run 2 vs Run 1, Run 3 vs Run 2 …)"
                : "All vs Baseline (Run 2 vs Run 1, Run 3 vs Run 1 …)");
            if (slaMs > 0) SetMeta(ws, genRow + 2, "SLA Threshold", $"{slaMs:0} ms");

            int tableRow = genRow + (slaMs > 0 ? 3 : 2) + 2;

            foreach (var pair in pairs)
            {
                tableRow = WriteSummarySection(ws, pair.Delta, $"{pair.LeftLabel}  →  {pair.RightLabel}", tableRow, slaMs);
                tableRow += 2;
            }

            ws.Cells.AutoFitColumns(14, 70);
        }

        private static int WriteSummarySection(
            ExcelWorksheet ws, List<ComparisonRecord> rows,
            string title, int startRow, double slaMs)
        {
            ws.Cells[startRow, 1].Value = title;
            ws.Cells[startRow, 1].Style.Font.Bold = true;
            ws.Cells[startRow, 1].Style.Font.Size = 13;
            ws.Cells[startRow, 1].Style.Font.Color.SetColor(ClrHdrFill);
            using (var bar = ws.Cells[startRow + 1, 1, startRow + 1, 20])
            {
                bar.Style.Fill.PatternType = ExcelFillStyle.Solid;
                bar.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xE0, 0xE7, 0xFF));
            }

            var matched = rows.Where(r => !r.OnlyInBaseline && !r.OnlyInCurrent).ToList();
            var regressions = matched.Where(r => r.DeltaAvgPct > RegressionThresholdPct).OrderByDescending(r => r.DeltaAvgPct).ToList();
            var improvements = matched.Where(r => r.DeltaAvgPct < -RegressionThresholdPct).OrderBy(r => r.DeltaAvgPct).ToList();

            int kpiRow = startRow + 2;
            WriteKpi(ws, kpiRow, 1, "Transactions Compared", matched.Count.ToString(), Color.FromArgb(0xDB, 0xEA, 0xFE), ClrNewText);
            WriteKpi(ws, kpiRow, 5, "Regressions (>10%)", regressions.Count.ToString(),
                regressions.Count > 0 ? ClrRegFill : ClrImpFill,
                regressions.Count > 0 ? ClrRegText : ClrImpText);
            WriteKpi(ws, kpiRow, 9, "Improvements (>10%)", improvements.Count.ToString(), ClrImpFill, ClrImpText);
            WriteKpi(ws, kpiRow, 13, "Only in Left", rows.Count(r => r.OnlyInBaseline).ToString(), ClrRemFill, ClrRemText);
            WriteKpi(ws, kpiRow, 17, "Only in Right", rows.Count(r => r.OnlyInCurrent).ToString(), ClrRemFill, ClrRemText);

            int tableRow = kpiRow + 4;

            if (regressions.Count > 0)
            {
                tableRow = WriteSectionLabel(ws, tableRow, "Top Regressions (by Avg)", ClrRegText);
                WriteTableHeader(ws, tableRow, new[] { "Transaction", "Left Avg (s)", "Right Avg (s)", "Delta (ms)", "Delta %" });
                tableRow++;
                foreach (var rec in regressions.Take(10))
                {
                    ws.Cells[tableRow, 1].Value = rec.TransactionName;
                    ws.Cells[tableRow, 2].Value = Math.Round(rec.BaseAvg / 1000.0, 3);
                    ws.Cells[tableRow, 3].Value = Math.Round(rec.CurAvg / 1000.0, 3);
                    ws.Cells[tableRow, 4].Value = Math.Round(rec.DeltaAvgMs, 0);
                    ws.Cells[tableRow, 5].Value = rec.DeltaAvgPct / 100.0;
                    ws.Cells[tableRow, 5].Style.Numberformat.Format = "+0.00%;-0.00%";
                    ws.Cells[tableRow, 5].Style.Font.Color.SetColor(ClrRegText);
                    ws.Cells[tableRow, 5].Style.Font.Bold = true;
                    ApplyRowFill(ws, tableRow, 1, 5, ClrRegFill);
                    tableRow++;
                }
                tableRow++;
            }

            if (slaMs > 0)
            {
                var slaBreaches = matched.Where(r => r.CurP90 > slaMs).OrderByDescending(r => r.CurP90).ToList();
                if (slaBreaches.Count > 0)
                {
                    tableRow = WriteSectionLabel(ws, tableRow, $"SLA Breaches — P90 > {slaMs:0} ms", ClrRemText);
                    WriteTableHeader(ws, tableRow, new[] { "Transaction", "Left P90 (s)", "Right P90 (s)", "Over SLA by (ms)" });
                    tableRow++;
                    foreach (var rec in slaBreaches)
                    {
                        ws.Cells[tableRow, 1].Value = rec.TransactionName;
                        ws.Cells[tableRow, 2].Value = Math.Round(rec.BaseP90 / 1000.0, 3);
                        ws.Cells[tableRow, 3].Value = Math.Round(rec.CurP90 / 1000.0, 3);
                        ws.Cells[tableRow, 4].Value = Math.Round(rec.CurP90 - slaMs, 0);
                        ws.Cells[tableRow, 4].Style.Font.Color.SetColor(ClrRemText);
                        ws.Cells[tableRow, 4].Style.Font.Bold = true;
                        ApplyRowFill(ws, tableRow, 1, 4, ClrRemFill);
                        tableRow++;
                    }
                }
            }
            return tableRow;
        }




        private static void WriteAvgSheet(
            ExcelPackage pkg, List<ComparisonRecord> rows, double slaMs,
            string prefix = "", string leftLabel = "", string rightLabel = "")
        {
            // Sort: A-Z by transaction name; one-sided records at the bottom
            var sorted = rows
                .OrderBy(r => r.OnlyInBaseline || r.OnlyInCurrent ? 1 : 0)
                .ThenBy(r => r.TransactionName, StringComparer.Ordinal)
                .ToList();

            var ws = pkg.Workbook.Worksheets.Add($"{prefix}Avg Comparison");
            string leftLbl = string.IsNullOrEmpty(leftLabel) ? "Baseline" : leftLabel;
            string rightLbl = string.IsNullOrEmpty(rightLabel) ? "Current" : rightLabel;

            var hdrs = BuildHdrList(
                "Transaction", "Status",
                $"{leftLbl} Avg (s)", $"{rightLbl} Avg (s)", "Δ Avg (ms)", "Δ Avg %",
                $"{leftLbl} Err %", $"{rightLbl} Err %", "Δ Err pts",
                $"{leftLbl} Samples", $"{rightLbl} Samples");

            if (slaMs > 0) hdrs.Add("SLA Breach (P90)");



            WriteSheetHeader(ws, hdrs.ToArray());



            int row = 2;

            foreach (var r in sorted)

            {

                string status = GetStatus(r, useP90: false);

                var (fill, _) = GetRowColors(status);

                int c = 1;



                ws.Cells[row, c++].Value = r.TransactionName;

                WriteStatusCell(ws, row, c++, status);



                ws.Cells[row, c++].Value = r.OnlyInCurrent ? (object)"—" : Math.Round(r.BaseAvg / 1000.0, 3);

                ws.Cells[row, c++].Value = r.OnlyInBaseline ? (object)"—" : Math.Round(r.CurAvg / 1000.0, 3);

                ws.Cells[row, c++].Value = BothPresent(r) ? (object)Math.Round(r.DeltaAvgMs, 0) : "—";



                var pctCell = ws.Cells[row, c++];

                if (BothPresent(r))

                {

                    pctCell.Value = r.DeltaAvgPct / 100.0;

                    pctCell.Style.Numberformat.Format = "+0.00%;-0.00%";

                    pctCell.Style.Font.Color.SetColor(r.DeltaAvgPct > 0 ? ClrRegText : ClrImpText);

                    pctCell.Style.Font.Bold = true;

                }

                else pctCell.Value = "—";



                WriteErrCols(ws, row, r, ref c);



                ws.Cells[row, c++].Value = r.OnlyInCurrent ? (object)"—" : r.BaseSamples;

                ws.Cells[row, c++].Value = r.OnlyInBaseline ? (object)"—" : r.CurSamples;



                if (slaMs > 0)

                {

                    bool breach = !r.OnlyInBaseline && r.CurP90 > slaMs;

                    ws.Cells[row, c].Value = breach ? "YES" : "";

                    if (breach) { ws.Cells[row, c].Style.Font.Color.SetColor(ClrRegText); ws.Cells[row, c].Style.Font.Bold = true; }

                    c++;

                }



                ApplyRowFill(ws, row, 1, hdrs.Count, fill);

                ws.Cells[row, 1].Style.Font.Color.SetColor(ClrStaText);

                row++;

            }



            FinaliseSheet(ws, pkg, row, hdrs.Count, $"Avg{prefix.Replace(" ", "")}Comparison");

        }



        // ── P90 Comparison sheet ──────────────────────────────────────────────



        private static void WriteP90Sheet(
            ExcelPackage pkg, List<ComparisonRecord> rows, double slaMs,
            string prefix = "", string leftLabel = "", string rightLabel = "")
        {
            // Sort: A-Z by transaction name; one-sided records at the bottom
            var sorted = rows
                .OrderBy(r => r.OnlyInBaseline || r.OnlyInCurrent ? 1 : 0)
                .ThenBy(r => r.TransactionName, StringComparer.Ordinal)
                .ToList();

            var ws = pkg.Workbook.Worksheets.Add($"{prefix}P90 Comparison");
            string leftLbl = string.IsNullOrEmpty(leftLabel) ? "Baseline" : leftLabel;
            string rightLbl = string.IsNullOrEmpty(rightLabel) ? "Current" : rightLabel;

            var hdrs = BuildHdrList(
                "Transaction", "Status",
                $"{leftLbl} P90 (s)", $"{rightLbl} P90 (s)", "Δ P90 (ms)", "Δ P90 %",
                $"{leftLbl} Err %", $"{rightLbl} Err %", "Δ Err pts",
                $"{leftLbl} Samples", $"{rightLbl} Samples");

            if (slaMs > 0) hdrs.Add("SLA Breach");



            WriteSheetHeader(ws, hdrs.ToArray());



            int row = 2;

            foreach (var r in sorted)

            {

                string status = GetStatus(r, useP90: true);

                var (fill, _) = GetRowColors(status);

                int c = 1;



                ws.Cells[row, c++].Value = r.TransactionName;

                WriteStatusCell(ws, row, c++, status);



                ws.Cells[row, c++].Value = r.OnlyInCurrent ? (object)"—" : Math.Round(r.BaseP90 / 1000.0, 3);

                ws.Cells[row, c++].Value = r.OnlyInBaseline ? (object)"—" : Math.Round(r.CurP90 / 1000.0, 3);

                ws.Cells[row, c++].Value = BothPresent(r) ? (object)Math.Round(r.DeltaP90Ms, 0) : "—";



                var pctCell = ws.Cells[row, c++];

                if (BothPresent(r))

                {

                    pctCell.Value = r.DeltaP90Pct / 100.0;

                    pctCell.Style.Numberformat.Format = "+0.00%;-0.00%";

                    pctCell.Style.Font.Color.SetColor(r.DeltaP90Pct > 0 ? ClrRegText : ClrImpText);

                    pctCell.Style.Font.Bold = true;

                }

                else pctCell.Value = "—";



                WriteErrCols(ws, row, r, ref c);



                ws.Cells[row, c++].Value = r.OnlyInCurrent ? (object)"—" : r.BaseSamples;

                ws.Cells[row, c++].Value = r.OnlyInBaseline ? (object)"—" : r.CurSamples;



                if (slaMs > 0)

                {

                    bool breach = !r.OnlyInBaseline && r.CurP90 > slaMs;

                    ws.Cells[row, c].Value = breach ? "YES" : "";

                    if (breach) { ws.Cells[row, c].Style.Font.Color.SetColor(ClrRegText); ws.Cells[row, c].Style.Font.Bold = true; }

                    c++;

                }



                ApplyRowFill(ws, row, 1, hdrs.Count, fill);

                ws.Cells[row, 1].Style.Font.Color.SetColor(ClrStaText);

                row++;

            }



            FinaliseSheet(ws, pkg, row, hdrs.Count, $"P90{prefix.Replace(" ", "")}Comparison");

        }



        // ── Raw data sheet ────────────────────────────────────────────────────



        private static void WriteRawSheet(

            ExcelPackage pkg, List<FlatRecord> records, string sheetName)

        {

            var ws = pkg.Workbook.Worksheets.Add(sheetName);

            var cols = new[] { "Transaction", "Samples", "Avg (s)", "Median (s)",

                               "P90 (s)", "Min (s)", "Max (s)", "Error %" };



            WriteSheetHeader(ws, cols);



            int row = 2;

            foreach (var r in records.OrderBy(x => x.Name, StringComparer.Ordinal))

            {

                ws.Cells[row, 1].Value = r.Name;

                ws.Cells[row, 2].Value = r.Samples;

                ws.Cells[row, 3].Value = Math.Round(r.Average / 1000.0, 3);

                ws.Cells[row, 4].Value = Math.Round(r.Median / 1000.0, 3);

                ws.Cells[row, 5].Value = Math.Round(r.P90 / 1000.0, 3);

                ws.Cells[row, 6].Value = Math.Round(r.Min / 1000.0, 3);

                ws.Cells[row, 7].Value = Math.Round(r.Max / 1000.0, 3);

                ws.Cells[row, 8].Value = r.Errors / 100.0;

                ws.Cells[row, 8].Style.Numberformat.Format = "0.00%";

                if (row % 2 == 0) ApplyRowFill(ws, row, 1, cols.Length, ClrAltRow);

                row++;

            }



            ws.Cells.AutoFitColumns();

            var table = ws.Tables.Add(ws.Cells[1, 1, row - 1, cols.Length],

                JTLFileProcessing.UniqueTableName(pkg, SanitiseTableName(sheetName + "Raw")));

            table.ShowHeader = true;

            table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;

        }



        // ── Sheet helpers ─────────────────────────────────────────────────────



        private static List<string> BuildHdrList(params string[] items) => new List<string>(items);



        private static void WriteSheetHeader(ExcelWorksheet ws, string[] cols)

        {

            for (int c = 0; c < cols.Length; c++)

                ws.Cells[1, c + 1].Value = cols[c];

            using var hdr = ws.Cells[1, 1, 1, cols.Length];

            hdr.Style.Font.Bold = true;

            hdr.Style.Fill.PatternType = ExcelFillStyle.Solid;

            hdr.Style.Fill.BackgroundColor.SetColor(ClrHdrFill);

            hdr.Style.Font.Color.SetColor(ClrHdrText);

        }



        private static void FinaliseSheet(

            ExcelWorksheet ws, ExcelPackage pkg, int nextRow, int colCount, string tableName)

        {

            ws.Cells.AutoFitColumns(12, 55);

            ws.View.FreezePanes(2, 1);

            if (nextRow > 2)

            {

                var safeName = SanitiseTableName(tableName);

                var table = ws.Tables.Add(ws.Cells[1, 1, nextRow - 1, colCount],

                    JTLFileProcessing.UniqueTableName(pkg, safeName));

                table.ShowHeader = true;

                table.TableStyle = OfficeOpenXml.Table.TableStyles.Light1;

            }

        }



        // Excel table names: letters, digits, underscores only; must start with letter/_; max 255 chars

        private static string SanitiseTableName(string name)

        {

            var sb = new System.Text.StringBuilder();

            foreach (char c in name)

            {

                if (char.IsLetterOrDigit(c) || c == '_') sb.Append(c);

                else sb.Append('_');

            }

            string result = sb.ToString();

            if (result.Length == 0 || char.IsDigit(result[0]))

                result = "T" + result;

            return result.Length > 200 ? result[..200] : result;

        }



        private static void WriteStatusCell(ExcelWorksheet ws, int row, int col, string status)

        {

            ws.Cells[row, col].Value = status;

            ws.Cells[row, col].Style.Font.Bold = true;

            ws.Cells[row, col].Style.Font.Color.SetColor(status switch

            {

                "Regression" => ClrRegText,

                "Improvement" => ClrImpText,

                "New" => ClrNewText,

                "Removed" => ClrRemText,

                _ => Color.FromArgb(0x6B, 0x72, 0x80)

            });

        }



        private static void WriteErrCols(

            ExcelWorksheet ws, int row, ComparisonRecord r, ref int c)

        {

            var eBase = ws.Cells[row, c++];

            if (!r.OnlyInCurrent) { eBase.Value = r.BaseErrors / 100.0; eBase.Style.Numberformat.Format = "0.00%"; }

            else eBase.Value = "—";



            var eCur = ws.Cells[row, c++];

            if (!r.OnlyInBaseline) { eCur.Value = r.CurErrors / 100.0; eCur.Style.Numberformat.Format = "0.00%"; }

            else eCur.Value = "—";



            ws.Cells[row, c++].Value = BothPresent(r) ? (object)Math.Round(r.DeltaErrorPct, 2) : "—";

        }



        private static void ApplyRowFill(

            ExcelWorksheet ws, int row, int fromCol, int toCol, Color fill)

        {

            ws.Cells[row, fromCol, row, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;

            ws.Cells[row, fromCol, row, toCol].Style.Fill.BackgroundColor.SetColor(fill);

        }



        private static (Color fill, Color text) GetRowColors(string status) => status switch

        {

            "Regression" => (ClrRegFill, ClrStaText),

            "Improvement" => (ClrImpFill, ClrStaText),

            "New" => (ClrNewFill, ClrStaText),

            "Removed" => (ClrRemFill, ClrStaText),

            _ => (ClrStaFill, ClrStaText)

        };



        private static string GetStatus(ComparisonRecord r, bool useP90)

        {

            if (r.OnlyInBaseline) return "Removed";

            if (r.OnlyInCurrent) return "New";

            double pct = useP90 ? r.DeltaP90Pct : r.DeltaAvgPct;

            return pct > RegressionThresholdPct ? "Regression"

                 : pct < -RegressionThresholdPct ? "Improvement"

                 : "Stable";

        }



        private static bool BothPresent(ComparisonRecord r)

            => !r.OnlyInBaseline && !r.OnlyInCurrent;



        // ── Summary-specific helpers ──────────────────────────────────────────



        private static void SetMeta(ExcelWorksheet ws, int row, string label, string value)

        {

            ws.Cells[row, 1].Value = label;

            ws.Cells[row, 1].Style.Font.Bold = true;

            ws.Cells[row, 1].Style.Font.Color.SetColor(Color.FromArgb(0x6B, 0x72, 0x80));

            ws.Cells[row, 2].Value = value;

        }



        private static int WriteSectionLabel(

            ExcelWorksheet ws, int row, string title, Color color)

        {

            ws.Cells[row, 1].Value = title;

            ws.Cells[row, 1].Style.Font.Bold = true;

            ws.Cells[row, 1].Style.Font.Size = 13;

            ws.Cells[row, 1].Style.Font.Color.SetColor(color);

            return row + 1;

        }



        private static void WriteTableHeader(

            ExcelWorksheet ws, int row, string[] cols)

        {

            for (int c = 0; c < cols.Length; c++)

                ws.Cells[row, c + 1].Value = cols[c];

            using var hdr = ws.Cells[row, 1, row, cols.Length];

            hdr.Style.Font.Bold = true;

            hdr.Style.Fill.PatternType = ExcelFillStyle.Solid;

            hdr.Style.Fill.BackgroundColor.SetColor(ClrHdrFill);

            hdr.Style.Font.Color.SetColor(ClrHdrText);

        }



        private static void WriteKpi(

            ExcelWorksheet ws, int row, int col,

            string label, string value, Color bgColor, Color textColor)

        {

            var lbl = ws.Cells[row, col, row, col + 2];

            lbl.Merge = true;

            lbl.Value = label;

            lbl.Style.Font.Size = 9;

            lbl.Style.Font.Bold = true;

            lbl.Style.Font.Color.SetColor(textColor);

            lbl.Style.Fill.PatternType = ExcelFillStyle.Solid;

            lbl.Style.Fill.BackgroundColor.SetColor(bgColor);

            lbl.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

            lbl.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;



            var val = ws.Cells[row + 1, col, row + 1, col + 2];

            val.Merge = true;

            val.Value = value;

            val.Style.Font.Size = 22;

            val.Style.Font.Bold = true;

            val.Style.Font.Color.SetColor(textColor);

            val.Style.Fill.PatternType = ExcelFillStyle.Solid;

            val.Style.Fill.BackgroundColor.SetColor(bgColor);

            val.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            val.Style.VerticalAlignment = ExcelVerticalAlignment.Center;



            ws.Cells[row, col, row + 1, col + 2].Style.Border

                .BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(0xD1, 0xD5, 0xDB));

        }



        // ── Parsing helpers ───────────────────────────────────────────────────



        private static int ParseInt(string v) =>

            int.TryParse(v.Trim(), out var i) ? i : 0;



        private static double ParseMs(string v) =>

            double.TryParse(v.Trim(), NumberStyles.Any,

                CultureInfo.InvariantCulture, out var d) ? d : 0;



        private static double ParsePct(string v) =>

            double.TryParse(v.Replace("%", "").Trim(), NumberStyles.Any,

                CultureInfo.InvariantCulture, out var d) ? d : 0;



        /// <summary>
        /// Quote-aware CSV line splitter.  Handles fields wrapped in
        /// double-quotes and escaped quotes ("").
        /// </summary>
        private static string[] SplitCsvLine(string line)
        {
            var fields = new List<string>();
            var sb = new System.Text.StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                        else inQuotes = false;
                    }
                    else sb.Append(c);
                }
                else
                {
                    if (c == '"') inQuotes = true;
                    else if (c == ',') { fields.Add(sb.ToString()); sb.Clear(); }
                    else sb.Append(c);
                }
            }
            fields.Add(sb.ToString());
            return fields.ToArray();
        }

        // ── Internal flat record ──────────────────────────────────────────────



        private class FlatRecord

        {

            public string Name { get; set; } = string.Empty;

            public int Samples { get; set; }

            public double Average { get; set; }

            public double Median { get; set; }

            public double P90 { get; set; }

            public double Min { get; set; }

            public double Max { get; set; }

            public double Errors { get; set; }

        }

    }

}
