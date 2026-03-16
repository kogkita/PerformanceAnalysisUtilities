using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Globalization;
using System.IO;

namespace TestApp
{
    /// <summary>
    /// Represents one aggregated row (either a Transaction or a Web Request)
    /// computed from raw JTL samples.
    /// </summary>
    public class JTLFileProcessingRecord
    {
        public string TransactionName { get; set; } = string.Empty;
        public int Samples { get; set; }
        public double Average { get; set; }       // raw milliseconds (converted to seconds on output)
        public double Median { get; set; }        // raw milliseconds
        public double P90 { get; set; }           // raw milliseconds
        public double P80 { get; set; }           // raw milliseconds
        public double P70 { get; set; }           // raw milliseconds
        public double Min { get; set; }           // raw milliseconds
        public double Max { get; set; }           // raw milliseconds
        public double ErrorPercent { get; set; }  // 0-100
    }

    public static class JTLFileProcessing
    {
        // ─────────────────────────────────────────────────────────────────────
        // Public entry points  (mirror ResponseTimeConverter's API shape)
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>
        /// Converts a single JTL file to an Excel workbook saved at
        /// <paramref name="excelPath"/>.
        /// </summary>
        public static void Convert(string jtlPath, string excelPath, bool includeCharts = true)
        {
            ExcelPackage.License.SetNonCommercialPersonal("JTL File Processing");

            var records = ParseJtl(jtlPath);

            // Sheet 1: A-Z by transaction name
            var recordsAZ = records
                .OrderBy(r => r.TransactionName, StringComparer.Ordinal)
                .ToList();

            // Charts: highest average response time first
            var recordsByAvg = records
                .OrderByDescending(r => r.Average)
                .ToList();

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
        /// If <paramref name="excelPath"/> already exists and contains a
        /// "JTL Results" sheet, returns <paramref name="records"/> reordered
        /// to match that sheet's row order.  Any records not found in the
        /// existing sheet are appended at the end (new transactions).
        /// </summary>
        /// <summary>
        /// Appends sheets for one JTL file into an existing package.
        /// When <paramref name="prefix"/> is non-null the sheet names are
        /// prefixed for clubbed / multi-file mode.
        /// NOTE: for the clubbed path with charts, call SaveAndPatchCharts
        /// after all files are appended.
        /// </summary>
        public static void AppendToPackage(
            ExcelPackage package,
            string jtlPath,
            string? prefix,
            bool includeCharts = true)
        {
            var records = ParseJtl(jtlPath);
            AppendToPackageWithRecords(package, records, prefix, includeCharts);
        }

        internal static void AppendToPackageWithRecords(
            ExcelPackage package,
            List<JTLFileProcessingRecord> records,
            string? prefix,
            bool includeCharts = true)
        {
            string dataName = prefix != null ? $"{prefix} \u2013 JTL Results" : "JTL Results";
            dataName = UniqueSheetName(package, dataName);

            var dataSheet = WriteResultsSheet(package, records, dataName);

            if (includeCharts && records.Count > 0)
            {
                // Register N EPPlus chart shells — one per transaction row.
                // AddMiniChartsAndSave (or InjectPendingCharts) overwrites them after SaveAs.
                for (int i = 0; i < records.Count; i++)
                {
                    int sheetRow = 2 + i; // 1-based, row 1 = header
                    var c = (OfficeOpenXml.Drawing.Chart.ExcelBarChart)
                        dataSheet.Drawings.AddChart($"Chart_{prefix}_{i}",
                            OfficeOpenXml.Drawing.Chart.eChartType.BarClustered);
                    c.SetPosition(sheetRow - 1, 0, 10, 0); // col K = index 10
                    c.SetSize(1, 1);
                }
                _pendingCharts[dataName] = records.ToList();
            }
        }

        // Pending chart data for clubbed mode (sheet name → records)
        internal static readonly Dictionary<string, List<JTLFileProcessingRecord>>
            _pendingCharts = new();

        /// <summary>
        /// Called by MainWindow after SaveAs in clubbed mode to inject all
        /// pending chart XMLs into the saved file.
        /// </summary>
        public static void InjectPendingCharts(string xlsxPath)
        {
            foreach (var kvp in _pendingCharts)
                JTLFileProcessingExcelCharts.InjectChartForSheet(xlsxPath, kvp.Key, kvp.Value);
            _pendingCharts.Clear();
        }

        // ─────────────────────────────────────────────────────────────────────
        // Sheet-name helpers  (same pattern as ResponseTimeConverter)
        // ─────────────────────────────────────────────────────────────────────

        internal static string UniqueSheetName(ExcelPackage pkg, string name)
        {
            if (name.Length > 31) name = name[..31];
            string candidate = name;
            int n = 2;
            while (pkg.Workbook.Worksheets.Any(
                ws => ws.Name.Equals(candidate, StringComparison.OrdinalIgnoreCase)))
            {
                candidate = $"{name[..Math.Min(name.Length, 28)]} {n++}";
            }
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

        // ─────────────────────────────────────────────────────────────────────
        // JTL parsing  →  ordered list of aggregated records
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>Public entry point for parsing — used by MainWindow for clubbed mode.</summary>
        public static List<JTLFileProcessingRecord> ParseJtlPublic(string jtlPath)
            => ParseJtl(jtlPath);

        internal static List<JTLFileProcessingRecord> ParseJtl(string jtlPath)
        {
            if (!File.Exists(jtlPath))
                throw new FileNotFoundException("JTL file not found", jtlPath);

            if (new FileInfo(jtlPath).Length == 0)
                throw new InvalidDataException("JTL file is empty or has no data rows.");

            // ── Column indices (resolved from the header line) ────────────────
            int idxTimestamp = -1;
            int idxElapsed = -1;
            int idxLabel = -1;
            int idxSuccess = -1;
            int idxUrl = -1;
            int idxRespMessage = -1;

            // Per-label accumulators
            var labelElapsed = new Dictionary<string, List<int>>();
            var labelErrors = new Dictionary<string, int>();
            var labelFirstTs = new Dictionary<string, long>();
            var labelIsTx = new Dictionary<string, bool>();

            bool headerParsed = false;

            // ── RFC-4180 CSV reader — handles quoted fields that contain commas ─
            // The responseMessage column for transaction rows is a quoted string:
            //   "Number of samples in transaction : 3, number of failing samples : 0"
            // A naive Split(',') shifts every subsequent column index by +1, which
            // corrupts the success and URL fields.  We parse properly here.
            using var reader = new StreamReader(jtlPath);
            string? line;
            while ((line = reader.ReadLine()) != null)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;

                var cols = SplitCsvLine(line);

                if (!headerParsed)
                {
                    // First non-blank line is the header
                    for (int c = 0; c < cols.Count; c++)
                    {
                        switch (cols[c].Trim())
                        {
                            case "timeStamp": idxTimestamp = c; break;
                            case "elapsed": idxElapsed = c; break;
                            case "label": idxLabel = c; break;
                            case "success": idxSuccess = c; break;
                            case "URL": idxUrl = c; break;
                            case "responseMessage": idxRespMessage = c; break;
                        }
                    }

                    if (idxElapsed < 0 || idxLabel < 0)
                        throw new InvalidDataException(
                            "JTL file is missing required columns (elapsed, label).");

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

                // Transaction rows: URL field is "null"/empty, OR responseMessage
                // contains "Number of samples in transaction".
                bool isTx = false;
                if (idxUrl >= 0 && idxUrl < cols.Count)
                {
                    string url = cols[idxUrl].Trim();
                    isTx = url.Equals("null", StringComparison.OrdinalIgnoreCase)
                        || url.Length == 0;
                }
                if (!isTx && idxRespMessage >= 0 && idxRespMessage < cols.Count)
                    isTx = cols[idxRespMessage].Contains("Number of samples in transaction");

                // Accumulate
                if (!labelElapsed.ContainsKey(label))
                {
                    labelElapsed[label] = new List<int>();
                    labelErrors[label] = 0;
                    labelFirstTs[label] = ts;
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

            // Sort by first-occurrence timestamp — reproduces JMeter's aggregate
            // report ordering (transaction followed by its child web requests, etc.)
            var orderedLabels = labelFirstTs.Keys
                .OrderBy(l => labelFirstTs[l])
                .ThenBy(l => l, StringComparer.Ordinal)
                .ToList();

            // Build records — transactions only, web requests excluded
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

        // ─────────────────────────────────────────────────────────────────────
        // RFC-4180 CSV line splitter
        // Handles fields enclosed in double-quotes that may contain commas.
        // ─────────────────────────────────────────────────────────────────────

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
                        // Escaped quote ("") inside a quoted field
                        if (i + 1 < line.Length && line[i + 1] == '"')
                        {
                            sb.Append('"');
                            i++; // skip second quote
                        }
                        else
                        {
                            inQuotes = false; // closing quote
                        }
                    }
                    else
                    {
                        sb.Append(c);
                    }
                }
                else
                {
                    if (c == '"')
                    {
                        inQuotes = true;
                    }
                    else if (c == ',')
                    {
                        fields.Add(sb.ToString());
                        sb.Clear();
                    }
                    else
                    {
                        sb.Append(c);
                    }
                }
            }

            fields.Add(sb.ToString()); // last field
            return fields;
        }

        // ─────────────────────────────────────────────────────────────────────
        // Stats helpers
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>
        /// Matches JMeter's percentile rounding: index = int(n × p/100 + 0.5) − 1.
        /// Uses standard "round half up" to pick the index, then does a direct
        /// array lookup (no interpolation), which is how JMeter's StatCalculator works.
        /// The remaining tiny differences vs JMeter's Aggregate Report for very small
        /// samples (n ≤ 5) are caused by JMeter's internal histogram-bucket
        /// approximation and are not reproducible from raw samples alone.
        /// </summary>
        private static double Percentile(List<int> data, double p)
        {
            var sorted = data.OrderBy(v => v).ToList();
            int n = sorted.Count;
            int idx = (int)(n * p / 100.0 + 0.5) - 1;
            idx = Math.Max(0, Math.Min(idx, n - 1));
            return sorted[idx];
        }

        // ─────────────────────────────────────────────────────────────────────
        // Excel sheet writer
        // ─────────────────────────────────────────────────────────────────────

        private static ExcelWorksheet WriteResultsSheet(
            ExcelPackage package,
            List<JTLFileProcessingRecord> records,
            string sheetName)
        {
            var sheet = package.Workbook.Worksheets.Add(sheetName);

            // ── Header row ────────────────────────────────────────────────────
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
                hdr.Style.Fill.PatternType = ExcelFillStyle.Solid;
                hdr.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            }

            // ── Data rows ─────────────────────────────────────────────────────
            int row = 2;
            foreach (var r in records)
            {
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

            // ── Wrap in an Excel Table ────────────────────────────────────────
            int totalRows = records.Count + 1;
            var tableRange = sheet.Cells[1, 1, totalRows, totalCols];
            var table = sheet.Tables.Add(tableRange, UniqueTableName(package, "JTLResults"));
            table.ShowHeader = true;
            table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;

            return sheet;
        }
    }
}