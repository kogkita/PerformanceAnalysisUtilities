using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Xml.Linq;

namespace TestApp
{
    public class JTLRecord
    {
        public string Label { get; set; } = string.Empty;
        public long Timestamp { get; set; }
        public int Elapsed { get; set; }
        public bool Success { get; set; }
        public int Bytes { get; set; }
        public int SentBytes { get; set; }
        public int GrpThreads { get; set; }
        public int AllThreads { get; set; }
        public int Latency { get; set; }
        public int Connect { get; set; }
        public string ResponseCode { get; set; } = string.Empty;
        public string ResponseMessage { get; set; } = string.Empty;
        public string ThreadName { get; set; } = string.Empty;
        public string DataType { get; set; } = string.Empty;
        public string FailureMessage { get; set; } = string.Empty;
    }

    public class JTLSummary
    {
        public string Label { get; set; } = string.Empty;
        public int Samples { get; set; }
        public double AverageMs { get; set; }
        public double MedianMs { get; set; }
        public double P90Ms { get; set; }
        public double P95Ms { get; set; }
        public double P99Ms { get; set; }
        public double MinMs { get; set; }
        public double MaxMs { get; set; }
        public double ErrorPercent { get; set; }
        public double ThroughputPerSec { get; set; }
        public double AvgBytesKB { get; set; }
    }

    public static class JTLFileProcessing
    {
        public static void Process(string jtlPath, string excelPath)
        {
            ExcelPackage.License.SetNonCommercialPersonal("JTL File Processing");
            using var package = new ExcelPackage();
            AppendToPackage(package, jtlPath, prefix: null);
            package.SaveAs(new FileInfo(excelPath));
        }

        /// <summary>
        /// Appends sheets for one JTL file into an existing package.
        /// When <paramref name="prefix"/> is non-null the sheet names are prefixed for clubbed mode.
        /// </summary>
        public static void AppendToPackage(ExcelPackage package, string jtlPath, string? prefix)
        {
            var records = ReadJtl(jtlPath);

            if (records.Count == 0)
                throw new InvalidDataException("No valid records found in the JTL file.");

            var summaries = BuildSummaries(records);

            string rawName = prefix != null ? $"{prefix} – Raw Data" : "Raw Data";
            string summaryName = prefix != null ? $"{prefix} – Summary" : "Summary";
            string chartName = prefix != null ? $"{prefix} – Latency Charts" : "Latency Charts";

            rawName = UniqueSheetName(package, rawName);
            summaryName = UniqueSheetName(package, summaryName);
            chartName = UniqueSheetName(package, chartName);

            WriteRawSheet(package, records, rawName);
            var summarySheet = WriteSummarySheet(package, summaries, summaryName);

            if (summaries.Count > 0)
                CreateChartSheet(package, summarySheet, summaries, chartName);
        }

        private static string UniqueSheetName(ExcelPackage pkg, string name)
        {
            if (name.Length > 31) name = name[..31];
            string candidate = name;
            int n = 2;
            while (pkg.Workbook.Worksheets.Any(ws => ws.Name.Equals(candidate, StringComparison.OrdinalIgnoreCase)))
                candidate = $"{name[..Math.Min(name.Length, 28)]} {n++}";
            return candidate;
        }

        // ── Readers ──────────────────────────────────────────────────────────

        private static List<JTLRecord> ReadJtl(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();

            return ext switch
            {
                ".xml" or ".jtl" when IsXml(path) => ReadXml(path),
                _ => ReadCsv(path)
            };
        }

        private static bool IsXml(string path)
        {
            try
            {
                using var sr = new StreamReader(path);
                var firstChar = (char)sr.Read();
                return firstChar == '<';
            }
            catch { return false; }
        }

        private static List<JTLRecord> ReadXml(string path)
        {
            var records = new List<JTLRecord>();
            var doc = XDocument.Load(path);

            foreach (var el in doc.Descendants("httpSample")
                              .Concat(doc.Descendants("sample")))
            {
                records.Add(new JTLRecord
                {
                    Label = (string?)el.Attribute("lb") ?? string.Empty,
                    Timestamp = ParseLong(el.Attribute("ts")?.Value),
                    Elapsed = ParseInt(el.Attribute("t")?.Value),
                    Success = (el.Attribute("s")?.Value ?? "true")
                              .Equals("true", StringComparison.OrdinalIgnoreCase),
                    Bytes = ParseInt(el.Attribute("by")?.Value),
                    SentBytes = ParseInt(el.Attribute("sby")?.Value),
                    GrpThreads = ParseInt(el.Attribute("ng")?.Value),
                    AllThreads = ParseInt(el.Attribute("na")?.Value),
                    Latency = ParseInt(el.Attribute("lt")?.Value),
                    Connect = ParseInt(el.Attribute("ct")?.Value),
                    ResponseCode = (string?)el.Attribute("rc") ?? string.Empty,
                    ResponseMessage = (string?)el.Attribute("rm") ?? string.Empty,
                    ThreadName = (string?)el.Attribute("tn") ?? string.Empty,
                    DataType = (string?)el.Attribute("dt") ?? string.Empty,
                    FailureMessage = el.Element("assertionResult")?.Element("failureMessage")?.Value ?? string.Empty
                });
            }

            return records;
        }

        private static List<JTLRecord> ReadCsv(string path)
        {
            var records = new List<JTLRecord>();
            var lines = File.ReadAllLines(path);

            if (lines.Length < 2)
                return records;

            var headers = lines[0].Split(',');

            int Idx(string name) => Array.IndexOf(headers, name);

            int iTs = Idx("timeStamp");
            int iElapsed = Idx("elapsed");
            int iLabel = Idx("label");
            int iRc = Idx("responseCode");
            int iRm = Idx("responseMessage");
            int iTn = Idx("threadName");
            int iDt = Idx("dataType");
            int iSuccess = Idx("success");
            int iBytes = Idx("bytes");
            int iSent = Idx("sentBytes");
            int iGrp = Idx("grpThreads");
            int iAll = Idx("allThreads");
            int iLat = Idx("Latency");
            int iConn = Idx("Connect");
            int iMsg = Idx("failureMessage");

            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;

                var v = SplitCsvLine(lines[i]);

                records.Add(new JTLRecord
                {
                    Timestamp = iTs >= 0 ? ParseLong(v[iTs]) : 0,
                    Elapsed = iElapsed >= 0 ? ParseInt(v[iElapsed]) : 0,
                    Label = iLabel >= 0 ? v[iLabel] : string.Empty,
                    ResponseCode = iRc >= 0 ? v[iRc] : string.Empty,
                    ResponseMessage = iRm >= 0 ? v[iRm] : string.Empty,
                    ThreadName = iTn >= 0 ? v[iTn] : string.Empty,
                    DataType = iDt >= 0 ? v[iDt] : string.Empty,
                    Success = iSuccess >= 0 && v[iSuccess].Equals("true", StringComparison.OrdinalIgnoreCase),
                    Bytes = iBytes >= 0 ? ParseInt(v[iBytes]) : 0,
                    SentBytes = iSent >= 0 ? ParseInt(v[iSent]) : 0,
                    GrpThreads = iGrp >= 0 ? ParseInt(v[iGrp]) : 0,
                    AllThreads = iAll >= 0 ? ParseInt(v[iAll]) : 0,
                    Latency = iLat >= 0 ? ParseInt(v[iLat]) : 0,
                    Connect = iConn >= 0 ? ParseInt(v[iConn]) : 0,
                    FailureMessage = iMsg >= 0 ? v[iMsg] : string.Empty
                });
            }

            return records;
        }

        // ── Aggregation ──────────────────────────────────────────────────────

        private static List<JTLSummary> BuildSummaries(List<JTLRecord> records)
        {
            var groups = records
                .GroupBy(r => r.Label)
                .OrderBy(g => g.Key);

            var summaries = new List<JTLSummary>();

            foreach (var g in groups)
            {
                var elapsed = g.Select(r => (double)r.Elapsed).OrderBy(x => x).ToList();
                int n = elapsed.Count;
                int errors = g.Count(r => !r.Success);

                double durationSec = 0;
                if (n > 1)
                {
                    long minTs = g.Min(r => r.Timestamp);
                    long maxTs = g.Max(r => r.Timestamp) + g.OrderByDescending(r => r.Timestamp).First().Elapsed;
                    durationSec = (maxTs - minTs) / 1000.0;
                }

                summaries.Add(new JTLSummary
                {
                    Label = g.Key,
                    Samples = n,
                    AverageMs = elapsed.Average(),
                    MedianMs = Percentile(elapsed, 50),
                    P90Ms = Percentile(elapsed, 90),
                    P95Ms = Percentile(elapsed, 95),
                    P99Ms = Percentile(elapsed, 99),
                    MinMs = elapsed.Min(),
                    MaxMs = elapsed.Max(),
                    ErrorPercent = n > 0 ? (errors / (double)n) * 100.0 : 0,
                    ThroughputPerSec = durationSec > 0 ? n / durationSec : 0,
                    AvgBytesKB = g.Average(r => r.Bytes) / 1024.0
                });
            }

            return summaries;
        }

        private static double Percentile(List<double> sorted, double p)
        {
            if (sorted.Count == 0) return 0;
            double rank = (p / 100.0) * (sorted.Count - 1);
            int lower = (int)rank;
            int upper = Math.Min(lower + 1, sorted.Count - 1);
            double frac = rank - lower;
            return sorted[lower] + frac * (sorted[upper] - sorted[lower]);
        }

        // ── Excel sheets ─────────────────────────────────────────────────────

        private static ExcelWorksheet WriteRawSheet(ExcelPackage package, List<JTLRecord> records, string sheetName = "Raw Data")
        {
            var sheet = package.Workbook.Worksheets.Add(sheetName);

            string[] headers =
            [
                "Label", "Timestamp", "Elapsed (ms)", "Response Code", "Response Message",
                "Thread Name", "Data Type", "Success", "Bytes", "Sent Bytes",
                "Group Threads", "All Threads", "Latency (ms)", "Connect (ms)", "Failure Message"
            ];

            for (int c = 0; c < headers.Length; c++)
                sheet.Cells[1, c + 1].Value = headers[c];

            using (var range = sheet.Cells[1, 1, 1, headers.Length])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.SteelBlue);
                range.Style.Font.Color.SetColor(Color.White);
            }

            int row = 2;
            foreach (var r in records)
            {
                sheet.Cells[row, 1].Value = r.Label;
                sheet.Cells[row, 2].Value = r.Timestamp;
                sheet.Cells[row, 3].Value = r.Elapsed;
                sheet.Cells[row, 4].Value = r.ResponseCode;
                sheet.Cells[row, 5].Value = r.ResponseMessage;
                sheet.Cells[row, 6].Value = r.ThreadName;
                sheet.Cells[row, 7].Value = r.DataType;
                sheet.Cells[row, 8].Value = r.Success ? "PASS" : "FAIL";
                sheet.Cells[row, 9].Value = r.Bytes;
                sheet.Cells[row, 10].Value = r.SentBytes;
                sheet.Cells[row, 11].Value = r.GrpThreads;
                sheet.Cells[row, 12].Value = r.AllThreads;
                sheet.Cells[row, 13].Value = r.Latency;
                sheet.Cells[row, 14].Value = r.Connect;
                sheet.Cells[row, 15].Value = r.FailureMessage;

                // Highlight failed rows
                if (!r.Success)
                {
                    using var rowRange = sheet.Cells[row, 1, row, headers.Length];
                    rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rowRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 50, 30, 30));
                    rowRange.Style.Font.Color.SetColor(Color.FromArgb(255, 248, 113, 113));
                }

                row++;
            }

            sheet.Cells.AutoFitColumns();
            return sheet;
        }

        private static ExcelWorksheet WriteSummarySheet(ExcelPackage package, List<JTLSummary> summaries, string sheetName = "Summary")
        {
            var sheet = package.Workbook.Worksheets.Add(sheetName);

            string[] headers =
            [
                "Transaction Name", "# Samples", "Average (s)", "Median (s)",
                "90th Pct (s)", "95th Pct (s)", "99th Pct (s)",
                "Min (s)", "Max (s)", "Error %", "Throughput (req/s)", "Avg KB/req"
            ];

            for (int c = 0; c < headers.Length; c++)
                sheet.Cells[1, c + 1].Value = headers[c];

            using (var range = sheet.Cells[1, 1, 1, headers.Length])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            }

            int row = 2;
            foreach (var s in summaries)
            {
                sheet.Cells[row, 1].Value = s.Label;
                sheet.Cells[row, 2].Value = s.Samples;
                sheet.Cells[row, 3].Value = Math.Round(s.AverageMs / 1000.0, 3);
                sheet.Cells[row, 4].Value = Math.Round(s.MedianMs / 1000.0, 3);
                sheet.Cells[row, 5].Value = Math.Round(s.P90Ms / 1000.0, 3);
                sheet.Cells[row, 6].Value = Math.Round(s.P95Ms / 1000.0, 3);
                sheet.Cells[row, 7].Value = Math.Round(s.P99Ms / 1000.0, 3);
                sheet.Cells[row, 8].Value = Math.Round(s.MinMs / 1000.0, 3);
                sheet.Cells[row, 9].Value = Math.Round(s.MaxMs / 1000.0, 3);

                var errCell = sheet.Cells[row, 10];
                errCell.Value = s.ErrorPercent / 100.0;
                errCell.Style.Numberformat.Format = "0.00%";

                // Highlight rows with errors
                if (s.ErrorPercent > 0)
                {
                    using var errRange = sheet.Cells[row, 10, row, 10];
                    errRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    errRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 220, 80, 80));
                    errRange.Style.Font.Color.SetColor(Color.White);
                    errRange.Style.Font.Bold = true;
                }

                sheet.Cells[row, 11].Value = Math.Round(s.ThroughputPerSec, 2);
                sheet.Cells[row, 12].Value = Math.Round(s.AvgBytesKB, 2);

                row++;
            }

            sheet.Cells.AutoFitColumns();

            // Wrap data in an Excel Table so sort order is reflected in the chart
            int totalRows = summaries.Count + 1;
            var tableRange = sheet.Cells[1, 1, totalRows, 12];
            var table = sheet.Tables.Add(tableRange, UniqueTableName(package, "JTLSummary"));
            table.ShowHeader = true;
            table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;

            return sheet;
        }

        private static void CreateChartSheet(
            ExcelPackage package,
            ExcelWorksheet summarySheet,
            List<JTLSummary> summaries,
            string sheetName = "Latency Charts")
        {
            var chartSheet = package.Workbook.Worksheets.Add(sheetName);

            var chart = chartSheet.Drawings.AddChart("JTLLatencyChart", eChartType.BarClustered);
            chart.Title.Text = "Latency Percentile Comparison";

            int recordCount = summaries.Count;
            int lastRow = recordCount + 1;

            var seriesDefs = new (int col, string label)[]
            {
                (3, "Average"),
                (5, "90th Pct"),
                (6, "95th Pct"),
                (7, "99th Pct")
            };

            foreach (var (col, label) in seriesDefs)
            {
                var series = chart.Series.Add(
                    summarySheet.Cells[2, col, lastRow, col],
                    summarySheet.Cells[2, 1, lastRow, 1]);
                series.Header = label;
            }

            // Compute outlier-resistant axis max
            // Collect percentile values in seconds (matching what's written to the sheet)
            var allValues = summaries
                .SelectMany(s => new[] { s.P90Ms / 1000.0, s.P95Ms / 1000.0, s.P99Ms / 1000.0 })
                .Where(v => v > 0)
                .OrderBy(v => v)
                .ToList();

            double? axisMax = null;
            if (allValues.Count > 0)
            {
                double p75 = allValues[(int)(allValues.Count * 0.75)];
                double hardMax = allValues[^1];
                if (hardMax > p75 * 3)
                    axisMax = Math.Ceiling(p75 * 1.5 * 10) / 10;
            }

            int chartHeight = Math.Max(500, recordCount * 40 + 100);
            chart.SetPosition(1, 0, 1, 0);
            chart.SetSize(900, chartHeight);

            FixBarChartAxisOrientation(chart, axisMax);
        }

        /// <summary>
        /// Directly patches the chart XML to produce a horizontal bar chart where:
        ///   - Category axis (transaction names) runs top-to-bottom (row 1 at top)
        ///   - Value axis (numbers) runs left-to-right with 0 on the left
        /// EPPlus's built-in Orientation/Crosses properties do not write the correct
        /// OOXML combination reliably, so we write the nodes ourselves.
        /// </summary>
        private static void FixBarChartAxisOrientation(ExcelChart chart, double? axisMax = null)
        {
            var xml = chart.ChartXml;
            var ns = new System.Xml.XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            // Category axis: reverse order top-to-bottom, value axis line at bottom,
            // labels on the LEFT ("low" = opposite end from where axis crosses at "max")
            var catAx = xml.SelectSingleNode("//c:catAx", ns);
            if (catAx != null)
            {
                SetOrCreateChildVal(xml, ns, catAx, "c:scaling/c:orientation", "maxMin");
                SetOrCreateChildVal(xml, ns, catAx, "c:crosses", "max");
                SetOrCreateChildVal(xml, ns, catAx, "c:tickLblPos", "low");
            }

            // Value axis: orientation normal left-to-right (0 on left, bars grow right).
            // crossesAt=0 → category axis line sits at value 0 (left edge).
            // tickLblPos=low → numbers at the bottom.
            var valAx = xml.SelectSingleNode("//c:valAx", ns);
            if (valAx != null)
            {
                SetOrCreateChildVal(xml, ns, valAx, "c:scaling/c:orientation", "minMax");
                if (axisMax.HasValue)
                    SetOrCreateChildVal(xml, ns, valAx, "c:scaling/c:max", axisMax.Value.ToString("G", System.Globalization.CultureInfo.InvariantCulture));
                SetOrCreateChildVal(xml, ns, valAx, "c:crossesAt", "0");
                SetOrCreateChildVal(xml, ns, valAx, "c:tickLblPos", "low");
            }
        }

        private static void SetOrCreateChildVal(
            System.Xml.XmlDocument xml,
            System.Xml.XmlNamespaceManager ns,
            System.Xml.XmlNode parent,
            string relPath,
            string val)
        {
            const string chartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            var parts = relPath.Split('/');
            var node = parent;

            foreach (var part in parts)
            {
                var child = node.SelectSingleNode(part, ns);
                if (child == null)
                {
                    var localName = part.Contains(':') ? part.Split(':')[1] : part;
                    child = xml.CreateElement("c", localName, chartNs);
                    node.AppendChild(child);
                }
                node = child;
            }

            var attr = node.Attributes?["val"];
            if (attr == null)
            {
                attr = xml.CreateAttribute("val");
                node.Attributes!.Append(attr);
            }
            attr.Value = val;
        }

        private static string UniqueTableName(ExcelPackage pkg, string name)
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

        // ── Helpers ──────────────────────────────────────────────────────────

        private static int ParseInt(string? value)
            => int.TryParse(value, out int r) ? r : 0;

        private static long ParseLong(string? value)
            => long.TryParse(value, out long r) ? r : 0;

        private static string[] SplitCsvLine(string line)
        {
            var result = new List<string>();
            bool inQuotes = false;
            var current = new System.Text.StringBuilder();

            foreach (char c in line)
            {
                if (c == '"') { inQuotes = !inQuotes; }
                else if (c == ',' && !inQuotes) { result.Add(current.ToString()); current.Clear(); }
                else { current.Append(c); }
            }

            result.Add(current.ToString());
            return result.ToArray();
        }
    }
}
