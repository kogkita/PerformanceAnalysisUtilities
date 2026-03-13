using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System.IO;

namespace TestApp
{
    public class ResponseTimeRecord
    {
        public string TransactionName { get; set; }
        public int Samples { get; set; }
        public double Average { get; set; }
        public double Median { get; set; }
        public Dictionary<string, double> Percentiles { get; set; } = new();
        public double Min { get; set; }
        public double Max { get; set; }
        public double ErrorPercent { get; set; }
    }

    public static class ResponseTimeConverter
    {
        // ─────────────────────────────────────────────────────────────────────
        // Public entry points
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>
        /// Converts a single CSV file to an Excel workbook saved at
        /// <paramref name="excelPath"/>.
        /// </summary>
        public static void Convert(string csvPath, string excelPath, bool includeCharts = true)
        {
            ExcelPackage.License.SetNonCommercialPersonal("Response Time Converter");

            using var package = new ExcelPackage();
            AppendToPackage(package, csvPath, prefix: null, includeCharts: includeCharts);
            package.SaveAs(new FileInfo(excelPath));
        }

        /// <summary>
        /// Appends sheets for one CSV into an existing package.
        /// When <paramref name="prefix"/> is non-null the sheet names are
        /// prefixed for clubbed / multi-file mode.
        /// </summary>
        public static void AppendToPackage(
            ExcelPackage package,
            string csvPath,
            string? prefix,
            bool includeCharts = true)
        {
            var (records, percentileHeaders) = ReadCsv(csvPath);

            string dataName = prefix != null ? $"{prefix} – Response Times" : "Response Times";
            string chartName = prefix != null ? $"{prefix} – Latency Charts" : "Latency Charts";

            dataName = UniqueSheetName(package, dataName);
            chartName = UniqueSheetName(package, chartName);

            var (dataSheet, percentileStartCol) =
                WriteResponseSheet(package, records, percentileHeaders, dataName);

            if (includeCharts && records.Count > 0)
            {
                ResponseTimeConverterExcelCharts.CreateChartSheet(
                    package,
                    dataSheet,
                    records,
                    percentileHeaders,
                    percentileStartCol,
                    chartName);
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        // Sheet-name / table-name helpers
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>
        /// Returns a sheet name that is unique within <paramref name="pkg"/>
        /// and within Excel's 31-character limit.
        /// </summary>
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

        /// <summary>
        /// Returns a table name that is unique across all worksheets in
        /// <paramref name="pkg"/> (Excel requires workbook-wide uniqueness).
        /// </summary>
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
        // CSV parsing
        // ─────────────────────────────────────────────────────────────────────

        private static (List<ResponseTimeRecord> records, List<string> percentileHeaders)
            ReadCsv(string csvPath)
        {
            var records = new List<ResponseTimeRecord>();
            var percentileHeaders = new List<string>();

            if (!File.Exists(csvPath))
                throw new FileNotFoundException("CSV file not found", csvPath);

            var lines = File.ReadAllLines(csvPath);
            var headers = lines[0].Split(',');

            // ── Column indices ────────────────────────────────────────────────
            int labelIndex = Array.IndexOf(headers, "Label");
            int sampleIndex = Array.IndexOf(headers, "# Samples");
            int avgIndex = Array.IndexOf(headers, "Average");
            int medianIndex = Array.IndexOf(headers, "Median");
            int minIndex = Array.IndexOf(headers, "Min");
            int maxIndex = Array.IndexOf(headers, "Max");
            int errIndex = Array.IndexOf(headers, "Error %");

            var percentileIndexes = new List<int>();
            for (int i = 0; i < headers.Length; i++)
            {
                if (headers[i].Contains("% Line"))
                {
                    percentileIndexes.Add(i);
                    percentileHeaders.Add(headers[i]);
                }
            }

            // ── Data rows ─────────────────────────────────────────────────────
            for (int i = 1; i < lines.Length; i++)
            {
                var values = lines[i].Split(',');

                // Skip the TOTAL summary row
                if (values[labelIndex].Trim()
                        .Equals("TOTAL", StringComparison.OrdinalIgnoreCase))
                    continue;

                var record = new ResponseTimeRecord
                {
                    TransactionName = values[labelIndex],
                    Samples = ParseInt(values[sampleIndex]),
                    Average = ToSeconds(values[avgIndex]),
                    Median = ToSeconds(values[medianIndex]),
                    Min = ToSeconds(values[minIndex]),
                    Max = ToSeconds(values[maxIndex]),
                    ErrorPercent = ParsePercent(values[errIndex])
                };

                foreach (var idx in percentileIndexes)
                    record.Percentiles[headers[idx]] = ToSeconds(values[idx]);

                records.Add(record);
            }

            return (records, percentileHeaders);
        }

        // ─────────────────────────────────────────────────────────────────────
        // Parsing helpers
        // ─────────────────────────────────────────────────────────────────────

        private static int ParseInt(string value) =>
            int.TryParse(value, out int result) ? result : 0;

        private static double ParsePercent(string value)
        {
            value = value.Replace("%", "");
            return double.TryParse(
                value,
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out double result) ? result : 0;
        }

        private static double ToSeconds(string ms) =>
            double.TryParse(
                ms,
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out double value) ? value / 1000 : 0;

        // ─────────────────────────────────────────────────────────────────────
        // Excel sheet writer
        // ─────────────────────────────────────────────────────────────────────

        private static (ExcelWorksheet sheet, int percentileStartCol)
            WriteResponseSheet(
                ExcelPackage package,
                List<ResponseTimeRecord> records,
                List<string> percentileHeaders,
                string sheetName = "Response Times")
        {
            var sheet = package.Workbook.Worksheets.Add(sheetName);
            int col = 1;

            // ── Header row ────────────────────────────────────────────────────
            sheet.Cells[1, col++].Value = "Transaction Name";
            sheet.Cells[1, col++].Value = "# Samples";
            sheet.Cells[1, col++].Value = "Average (Seconds)";
            sheet.Cells[1, col++].Value = "Median (Seconds)";

            int percentileStartCol = col; // capture BEFORE writing percentile headers

            foreach (var p in percentileHeaders)
                sheet.Cells[1, col++].Value = p.Replace("% Line", " Percentile (Seconds)");

            sheet.Cells[1, col++].Value = "Min (Seconds)";
            sheet.Cells[1, col++].Value = "Max (Seconds)";
            sheet.Cells[1, col++].Value = "Error %";

            // Style the header row
            using (var range = sheet.Cells[1, 1, 1, col - 1])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            }

            // ── Data rows ─────────────────────────────────────────────────────
            int row = 2;
            foreach (var r in records)
            {
                col = 1;

                sheet.Cells[row, col++].Value = r.TransactionName;
                sheet.Cells[row, col++].Value = r.Samples;
                sheet.Cells[row, col++].Value = r.Average;
                sheet.Cells[row, col++].Value = r.Median;

                foreach (var p in percentileHeaders)
                    sheet.Cells[row, col++].Value = r.Percentiles[p];

                sheet.Cells[row, col++].Value = r.Min;
                sheet.Cells[row, col++].Value = r.Max;

                var errorCell = sheet.Cells[row, col++];
                errorCell.Value = r.ErrorPercent / 100.0;
                errorCell.Style.Numberformat.Format = "0.00%";

                row++;
            }

            sheet.Cells.AutoFitColumns();

            // ── Wrap in an Excel Table ────────────────────────────────────────
            int totalRows = records.Count + 1; // header + data rows
            int totalCols = col - 1;
            var tableRange = sheet.Cells[1, 1, totalRows, totalCols];
            var table = sheet.Tables.Add(tableRange, UniqueTableName(package, "ResponseTimes"));
            table.ShowHeader = true;
            table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;

            return (sheet, percentileStartCol);
        }
    }
}