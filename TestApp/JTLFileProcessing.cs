using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Globalization;
using System.IO;

namespace TestApp
{
    public class JTLFileProcessingRecord
    {
        public string TransactionName { get; set; } = string.Empty;
        public int Samples { get; set; }
        public double Average { get; set; }       // raw milliseconds
        public double Median { get; set; }
        public double P90 { get; set; }
        public double P80 { get; set; }
        public double P70 { get; set; }
        public double Min { get; set; }
        public double Max { get; set; }
        public double ErrorPercent { get; set; }  // 0-100
    }

    public static class JTLFileProcessing
    {
        // ── Public entry points ───────────────────────────────────────────────

        public static void Convert(string jtlPath, string excelPath, bool includeCharts = true)
        {
            ExcelPackage.License.SetNonCommercialPersonal("JTL File Processing");

            var records = ParseJtl(jtlPath);

            // Data sheet: A-Z by transaction name
            var recordsAZ = records.OrderBy(r => r.TransactionName, StringComparer.Ordinal).ToList();
            // Charts: slowest average first
            var recordsByAvg = records.OrderByDescending(r => r.Average).ToList();

            using var package = new ExcelPackage();
            string dataName = UniqueSheetName(package, "JTL Results");
            var dataSheet = WriteResultsSheet(package, recordsAZ, dataName);

            if (includeCharts && records.Count > 0)
                JTLFileProcessingExcelCharts.AddMiniChartsAndSave(
                    package, dataSheet, recordsByAvg, excelPath);
            else
                package.SaveAs(new FileInfo(excelPath));
        }

        /// <summary>
        /// Appends sheets for one JTL file into an existing package.
        /// Call <see cref="InjectPendingCharts"/> after SaveAs to inject chart XML.
        /// </summary>
        public static void AppendToPackage(
            ExcelPackage package,
            string jtlPath,
            string? prefix,
            bool includeCharts = true)
        {
            var records = ParseJtl(jtlPath);

            string dataName = prefix != null ? $"{prefix} \u2013 JTL Results" : "JTL Results";
            string chartName = prefix != null ? $"{prefix} \u2013 JTL Charts" : "JTL Charts";
            dataName = UniqueSheetName(package, dataName);
            chartName = UniqueSheetName(package, chartName);

            var dataSheet = WriteResultsSheet(package, records.OrderBy(r => r.TransactionName, StringComparer.Ordinal).ToList(), dataName);

            if (includeCharts && records.Count > 0)
            {
                // Sort for charts: slowest average first
                var byAvg = records.OrderByDescending(r => r.Average).ToList();
                _pendingCharts[chartName] = byAvg;

                // Build the chart sheet structure (EPPlus shells only — XML injected after SaveAs)
                JTLFileProcessingExcelCharts.BuildChartSheetShells(package, chartName, byAvg);
            }
        }

        // Pending chart data for clubbed mode (chart sheet name → avg-sorted records)
        internal static readonly Dictionary<string, List<JTLFileProcessingRecord>>
            _pendingCharts = new();

        /// <summary>
        /// Clears any stale pending chart data.  Should be called at the
        /// start of a clubbed-mode run to avoid leaking data from a previous
        /// run that may have failed before <see cref="InjectPendingCharts"/>
        /// was reached.
        /// </summary>
        public static void ClearPendingCharts() => _pendingCharts.Clear();

        /// <summary>Called by MainWindow after SaveAs in clubbed mode.</summary>
        public static void InjectPendingCharts(string xlsxPath)
        {
            try
            {
                foreach (var kvp in _pendingCharts)
                    JTLFileProcessingExcelCharts.InjectChartsForSheet(xlsxPath, kvp.Key, kvp.Value);
            }
            finally
            {
                _pendingCharts.Clear();
            }
        }

        // ── Sheet-name helpers ────────────────────────────────────────────────

        internal static string UniqueSheetName(ExcelPackage pkg, string name)
        {
            if (name.Length > 31) name = name[..31];
            string candidate = name;
            int n = 2;
            while (pkg.Workbook.Worksheets.Any(
                ws => ws.Name.Equals(candidate, StringComparison.OrdinalIgnoreCase)))
                candidate = $"{name[..Math.Min(name.Length, 28)]} {n++}";
            return candidate;
        }

        internal static string UniqueTableName(ExcelPackage pkg, string name)
        {
            var existing = pkg.Workbook.Worksheets
                .SelectMany(ws => ws.Tables.Select(t => t.Name))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            string candidate = name;
            int n = 2;
            while (existing.Contains(candidate))
                candidate = $"{name}{n++}";
            return candidate;
        }

        // ── JTL parsing ───────────────────────────────────────────────────────

        internal static List<JTLFileProcessingRecord> ParseJtl(string jtlPath)
        {
            if (!File.Exists(jtlPath))
                throw new FileNotFoundException("JTL file not found", jtlPath);
            if (new FileInfo(jtlPath).Length == 0)
                throw new InvalidDataException("JTL file is empty or has no data rows.");

            int idxTimestamp = -1, idxElapsed = -1, idxLabel = -1;
            int idxSuccess = -1, idxUrl = -1, idxRespMessage = -1;

            var labelElapsed = new Dictionary<string, List<int>>();
            var labelErrors = new Dictionary<string, int>();
            var labelFirstTs = new Dictionary<string, long>();
            var labelIsTx = new Dictionary<string, bool>();
            bool headerParsed = false;

            using var reader = new StreamReader(jtlPath);
            string? line;
            while ((line = reader.ReadLine()) != null)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                var cols = SplitCsvLine(line);

                if (!headerParsed)
                {
                    for (int c = 0; c < cols.Count; c++)
                        switch (cols[c].Trim())
                        {
                            case "timeStamp": idxTimestamp = c; break;
                            case "elapsed": idxElapsed = c; break;
                            case "label": idxLabel = c; break;
                            case "success": idxSuccess = c; break;
                            case "URL": idxUrl = c; break;
                            case "responseMessage": idxRespMessage = c; break;
                        }
                    if (idxElapsed < 0 || idxLabel < 0)
                        throw new InvalidDataException("JTL file is missing required columns (elapsed, label).");
                    headerParsed = true;
                    continue;
                }

                if (cols.Count <= idxLabel) continue;
                string label = cols[idxLabel].Trim();
                if (string.IsNullOrEmpty(label)) continue;
                if (!int.TryParse(cols[idxElapsed].Trim(), out int elapsed)) continue;

                long ts = idxTimestamp >= 0 && idxTimestamp < cols.Count &&
                          long.TryParse(cols[idxTimestamp].Trim(), out long t) ? t : 0;
                bool success = idxSuccess < 0 || idxSuccess >= cols.Count ||
                    cols[idxSuccess].Trim().Equals("true", StringComparison.OrdinalIgnoreCase);

                bool isTx = false;
                if (idxUrl >= 0 && idxUrl < cols.Count)
                {
                    string url = cols[idxUrl].Trim();
                    isTx = url.Equals("null", StringComparison.OrdinalIgnoreCase) || url.Length == 0;
                }
                if (!isTx && idxRespMessage >= 0 && idxRespMessage < cols.Count)
                    isTx = cols[idxRespMessage].Contains("Number of samples in transaction");

                if (!labelElapsed.ContainsKey(label))
                {
                    labelElapsed[label] = new List<int>();
                    labelErrors[label] = 0;
                    labelFirstTs[label] = ts > 0 ? ts : long.MaxValue;
                    labelIsTx[label] = isTx;
                }
                else
                {
                    if (ts > 0 && ts < labelFirstTs[label]) labelFirstTs[label] = ts;
                    if (isTx) labelIsTx[label] = true;
                }
                labelElapsed[label].Add(elapsed);
                if (!success) labelErrors[label]++;
            }

            if (!headerParsed)
                throw new InvalidDataException("JTL file is empty or has no data rows.");

            var orderedLabels = labelFirstTs.Keys
                .OrderBy(l => labelFirstTs[l])
                .ThenBy(l => l, StringComparer.Ordinal)
                .ToList();

            var records = new List<JTLFileProcessingRecord>();
            foreach (var label in orderedLabels)
            {
                if (label.Equals("TOTAL", StringComparison.OrdinalIgnoreCase)) continue;
                if (!labelIsTx.TryGetValue(label, out bool isTxLabel) || !isTxLabel) continue;

                var data = labelElapsed[label];
                int samples = data.Count;
                double errPct = samples > 0 ? labelErrors[label] * 100.0 / samples : 0;

                records.Add(new JTLFileProcessingRecord
                {
                    TransactionName = label,
                    Samples = samples,
                    Average = data.Average(),
                    Median = Percentile(data, 50),
                    P90 = Percentile(data, 90),
                    P80 = Percentile(data, 80),
                    P70 = Percentile(data, 70),
                    Min = data.Min(),
                    Max = data.Max(),
                    ErrorPercent = errPct
                });
            }
            return records;
        }

        private static List<string> SplitCsvLine(string line)
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
            return fields;
        }

        private static double Percentile(List<int> data, double p)
        {
            var sorted = data.OrderBy(v => v).ToList();
            int n = sorted.Count;
            int idx = (int)(n * p / 100.0 + 0.5) - 1;
            idx = Math.Max(0, Math.Min(idx, n - 1));
            return sorted[idx];
        }

        // ── Excel sheet writer ────────────────────────────────────────────────

        private static ExcelWorksheet WriteResultsSheet(
            ExcelPackage package,
            List<JTLFileProcessingRecord> records,
            string sheetName)
        {
            var sheet = package.Workbook.Worksheets.Add(sheetName);

            sheet.Cells[1, 1].Value = "Label";
            sheet.Cells[1, 2].Value = "# Samples";
            sheet.Cells[1, 3].Value = "Average (Seconds)";
            sheet.Cells[1, 4].Value = "Median (Seconds)";
            sheet.Cells[1, 5].Value = "90% Line (Seconds)";
            sheet.Cells[1, 6].Value = "80% Line (Seconds)";
            sheet.Cells[1, 7].Value = "70% Line (Seconds)";
            sheet.Cells[1, 8].Value = "Min (Seconds)";
            sheet.Cells[1, 9].Value = "Max (Seconds)";
            sheet.Cells[1, 10].Value = "Error %";
            int totalCols = 10;

            using (var hdr = sheet.Cells[1, 1, 1, totalCols])
            {
                hdr.Style.Font.Bold = true;
                hdr.Style.Font.Color.SetColor(System.Drawing.Color.White);
                hdr.Style.Fill.PatternType = ExcelFillStyle.Solid;
                hdr.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0x1E, 0x40, 0xAF)); // dark blue
            }

            int row = 2;
            foreach (var r in records)
            {
                // Alternating row colour: white / very light grey
                if (row % 2 == 0)
                {
                    sheet.Cells[row, 1, row, totalCols].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[row, 1, row, totalCols].Style.Fill.BackgroundColor
                        .SetColor(System.Drawing.Color.FromArgb(0xF3, 0xF4, 0xF6));
                }

                sheet.Cells[row, 1].Value = r.TransactionName;
                sheet.Cells[row, 2].Value = r.Samples;
                sheet.Cells[row, 3].Value = Math.Round(r.Average / 1000.0, 3);
                sheet.Cells[row, 4].Value = Math.Round(r.Median / 1000.0, 3);
                sheet.Cells[row, 5].Value = Math.Round(r.P90 / 1000.0, 3);
                sheet.Cells[row, 6].Value = Math.Round(r.P80 / 1000.0, 3);
                sheet.Cells[row, 7].Value = Math.Round(r.P70 / 1000.0, 3);
                sheet.Cells[row, 8].Value = Math.Round(r.Min / 1000.0, 3);
                sheet.Cells[row, 9].Value = Math.Round(r.Max / 1000.0, 3);
                var errCell = sheet.Cells[row, 10];
                errCell.Value = r.ErrorPercent / 100.0;
                errCell.Style.Numberformat.Format = "0.00%";
                row++;
            }

            sheet.Cells.AutoFitColumns();
            if (records.Count > 0)
            {
                var tableRange = sheet.Cells[1, 1, records.Count + 1, totalCols];
                var table = sheet.Tables.Add(tableRange, UniqueTableName(package, "JTLResults"));
                table.ShowHeader = true;
                table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
            }
            return sheet;
        }
    }
}
