using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
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
        public static void Convert(string csvPath, string excelPath)
        {
            ExcelPackage.License.SetNonCommercialPersonal("Response Time Converter");

            using var package = new ExcelPackage();
            AppendToPackage(package, csvPath, prefix: null);
            package.SaveAs(new FileInfo(excelPath));
        }

        /// <summary>
        /// Appends sheets for one CSV into an existing package.
        /// When <paramref name="prefix"/> is non-null the sheet names are prefixed for clubbed mode.
        /// </summary>
        public static void AppendToPackage(ExcelPackage package, string csvPath, string? prefix)
        {
            var (records, percentileHeaders) = ReadCsv(csvPath);

            string dataName = prefix != null ? $"{prefix} – Response Times" : "Response Times";
            string chartName = prefix != null ? $"{prefix} – Latency Charts" : "Latency Charts";

            // Ensure unique names inside the same workbook
            dataName = UniqueSheetName(package, dataName);
            chartName = UniqueSheetName(package, chartName);

            var (dataSheet, percentileStartCol) = WriteResponseSheet(package, records, percentileHeaders, dataName);

            if (records.Count > 0)
                CreateChartSheet(package, dataSheet, records, percentileHeaders, percentileStartCol, chartName);
        }

        private static string UniqueSheetName(ExcelPackage pkg, string name)
        {
            // Excel sheet names max 31 chars
            if (name.Length > 31) name = name[..31];
            string candidate = name;
            int n = 2;
            while (pkg.Workbook.Worksheets.Any(ws => ws.Name.Equals(candidate, StringComparison.OrdinalIgnoreCase)))
                candidate = $"{name[..Math.Min(name.Length, 28)]} {n++}";
            return candidate;
        }

        private static (List<ResponseTimeRecord>, List<string>) ReadCsv(string csvPath)
        {
            var records = new List<ResponseTimeRecord>();
            var percentileHeaders = new List<string>();

            if (!File.Exists(csvPath))
                throw new FileNotFoundException("CSV file not found", csvPath);

            var lines = File.ReadAllLines(csvPath);
            var headers = lines[0].Split(',');

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

            for (int i = 1; i < lines.Length; i++)
            {
                var values = lines[i].Split(',');

                if (values[labelIndex].Trim().Equals("TOTAL", StringComparison.OrdinalIgnoreCase))
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
                {
                    record.Percentiles[headers[idx]] = ToSeconds(values[idx]);
                }

                records.Add(record);
            }

            return (records, percentileHeaders);
        }

        private static int ParseInt(string value)
        {
            return int.TryParse(value, out int result) ? result : 0;
        }

        private static double ParsePercent(string value)
        {
            value = value.Replace("%", "");
            return double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double result) ? result : 0;
        }

        private static double ToSeconds(string ms)
        {
            return double.TryParse(ms, NumberStyles.Any, CultureInfo.InvariantCulture, out double value)
                ? value / 1000
                : 0;
        }

        private static (ExcelWorksheet sheet, int percentileStartCol) WriteResponseSheet(
            ExcelPackage package,
            List<ResponseTimeRecord> records,
            List<string> percentileHeaders,
            string sheetName = "Response Times")
        {
            var sheet = package.Workbook.Worksheets.Add(sheetName);

            int col = 1;

            sheet.Cells[1, col++].Value = "Transaction Name";
            sheet.Cells[1, col++].Value = "# Samples";
            sheet.Cells[1, col++].Value = "Average (Seconds)";
            sheet.Cells[1, col++].Value = "Median (Seconds)";

            int percentileStartCol = col; // record BEFORE writing percentiles

            foreach (var p in percentileHeaders)
            {
                sheet.Cells[1, col++].Value = p.Replace("% Line", " Percentile (Seconds)");
            }

            sheet.Cells[1, col++].Value = "Min (Seconds)";
            sheet.Cells[1, col++].Value = "Max (Seconds)";
            sheet.Cells[1, col++].Value = "Error %";

            using (var range = sheet.Cells[1, 1, 1, col - 1])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            }

            int row = 2;

            foreach (var r in records)
            {
                col = 1;

                sheet.Cells[row, col++].Value = r.TransactionName;
                sheet.Cells[row, col++].Value = r.Samples;
                sheet.Cells[row, col++].Value = r.Average;
                sheet.Cells[row, col++].Value = r.Median;

                foreach (var p in percentileHeaders)
                {
                    sheet.Cells[row, col++].Value = r.Percentiles[p];
                }

                sheet.Cells[row, col++].Value = r.Min;
                sheet.Cells[row, col++].Value = r.Max;

                var errorCell = sheet.Cells[row, col++];
                errorCell.Value = r.ErrorPercent / 100.0;
                errorCell.Style.Numberformat.Format = "0.00%";

                row++;
            }

            sheet.Cells.AutoFitColumns();

            // Wrap data in an Excel Table so sort order is reflected in the chart
            int totalRows = records.Count + 1; // header + data
            int totalCols = col - 1;
            var tableRange = sheet.Cells[1, 1, totalRows, totalCols];
            var table = sheet.Tables.Add(tableRange, UniqueTableName(package, "ResponseTimes"));
            table.ShowHeader = true;
            table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;

            return (sheet, percentileStartCol);
        }

        private static string UniqueTableName(ExcelPackage pkg, string name)
        {
            // Table names must be unique across the whole workbook
            var existing = pkg.Workbook.Worksheets
                .SelectMany(ws => ws.Tables.Select(t => t.Name))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            string candidate = name;
            int n = 2;
            while (existing.Contains(candidate))
                candidate = $"{name}{n++}";
            return candidate;
        }

        private static void CreateChartSheet(
            ExcelPackage package,
            ExcelWorksheet dataSheet,
            List<ResponseTimeRecord> records,
            List<string> percentileHeaders,
            int percentileStartColumn,
            string sheetName = "Latency Charts")
        {
            var chartSheet = package.Workbook.Worksheets.Add(sheetName);

            var chart = chartSheet.Drawings.AddChart("LatencyChart", eChartType.BarClustered);
            chart.Title.Text = "Latency Percentile Comparison";

            int recordCount = records.Count;
            int lastRow = recordCount + 1;

            for (int i = 0; i < percentileHeaders.Count; i++)
            {
                int col = percentileStartColumn + i;
                var series = chart.Series.Add(
                    dataSheet.Cells[2, col, lastRow, col],
                    dataSheet.Cells[2, 1, lastRow, 1]);
                series.Header = percentileHeaders[i].Replace("% Line", "");
            }

            // Compute an outlier-resistant axis max:
            // Collect all percentile values, sort them, take the 75th percentile of those values,
            // then use 1.3× that as the chart max. This prevents one extreme outlier from
            // compressing all other bars into invisibility.
            var allValues = records
                .SelectMany(r => r.Percentiles.Values)
                .Where(v => v > 0)
                .OrderBy(v => v)
                .ToList();

            double? axisMax = null;
            if (allValues.Count > 0)
            {
                double p75 = allValues[(int)(allValues.Count * 0.75)];
                double hardMax = allValues[^1];
                // Only cap the axis if the outlier is more than 3× the p75 value
                if (hardMax > p75 * 3)
                    axisMax = Math.Ceiling(p75 * 1.5 * 10) / 10; // round up to 1 decimal
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

            // ── Category axis (catAx) ──────────────────────────────────────────
            // <c:scaling><c:orientation val="maxMin"/></c:scaling>  → reverses top-to-bottom
            // <c:crosses val="max"/>  → value axis line sits at the bottom of the reversed axis
            // <c:tickLblPos val="low"/> → labels appear on the LEFT side (low = away from crossing)
            var catAx = xml.SelectSingleNode("//c:catAx", ns);
            if (catAx != null)
            {
                SetOrCreateChildVal(xml, ns, catAx, "c:scaling/c:orientation", "maxMin");
                SetOrCreateChildVal(xml, ns, catAx, "c:crosses", "max");
                SetOrCreateChildVal(xml, ns, catAx, "c:tickLblPos", "low");
            }

            // ── Value axis (valAx) ────────────────────────────────────────────
            // orientation=minMax → 0 on left, max on right (bars grow leftward from 0)
            // crossesAt=0 → category axis intersects at value 0 (the left edge)
            // tickLblPos=low → number labels at the bottom
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
            string relPath,          // e.g. "c:scaling/c:orientation"
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
                    // Strip prefix to get local name
                    var localName = part.Contains(':') ? part.Split(':')[1] : part;
                    child = xml.CreateElement("c", localName, chartNs);
                    node.AppendChild(child);
                }
                node = child;
            }

            // node is now the leaf element — set/overwrite the val attribute
            if (node.Attributes == null) return;
            var attr = node.Attributes["val"];
            if (attr == null)
            {
                attr = xml.CreateAttribute("val");
                node.Attributes.Append(attr);
            }
            attr.Value = val;
        }
    }
}