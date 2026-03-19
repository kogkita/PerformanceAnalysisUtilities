using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System.IO;

namespace TestApp
{
    /// <summary>
    /// Reads one or more CSV files produced by relog.exe and generates an Excel workbook
    /// containing one line chart per key counter:
    ///   • Available Memory (MB)
    ///   • % CPU
    ///   • Pages Input/sec
    ///   • Disk Queue Length
    ///   • Network Interface (Mbps)
    ///   • % Disk Busy  (= 100 − % Idle Time)
    /// Each CSV = one server. The server label is derived from the file name.
    /// </summary>
    public static class BLGGraphProducer
    {
        // ── Chart definitions ─────────────────────────────────────────────────

        private record ChartDef(
            string Title,           // chart title shown in Excel
            string SearchKeyword,   // substring to find the counter header row
            bool ConvertToMbps,   // Network: bytes/s → Mbps  (÷1 000 000 × 8)
            bool InvertIdle       // Disk Busy: 100 − % Idle
        );

        private static readonly ChartDef[] Charts =
        {
            new("Available Memory (MB)",    "Available MBytes",          false, false),
            new("CPU Utilization",                    "Processor(_Total)",         false, false),
            new("Pages Input/sec",          "Pages Input/sec",           false, false),
            new("Disk Queue Length",        "Current Disk Queue Length", false, false),
            new("Network Interface (Mbps)", "Network Interface",         true,  false),
            new("Disk Busy %",              "% Idle Time",               false, true),
        };

        // ── Public entry point ────────────────────────────────────────────────

        /// <summary>
        /// Produces an Excel workbook with line charts from the supplied CSV paths.
        /// </summary>
        /// <param name="csvPaths">Ordered list of relog CSV files (one per server).</param>
        /// <param name="outputXlsxPath">Where to save the resulting workbook.</param>
        /// <param name="serverLabels">
        ///   Optional display names for each CSV. If null or shorter than csvPaths,
        ///   the file name (without extension) is used as a fallback.
        /// </param>
        public static void ProduceGraphs(
            IList<string> csvPaths,
            string outputXlsxPath,
            IList<string>? serverLabels = null)
        {
            if (csvPaths == null || csvPaths.Count == 0)
                throw new ArgumentException("At least one CSV file is required.");

            ExcelPackage.License.SetNonCommercialPersonal("BLG Graph Producer");

            // Load all CSVs into memory
            var servers = csvPaths
                .Select((path, idx) => new ServerData(
                    Label: (serverLabels != null && idx < serverLabels.Count && !string.IsNullOrWhiteSpace(serverLabels[idx]))
                           ? serverLabels[idx]
                           : Path.GetFileNameWithoutExtension(path),
                    Rows: ParseCsv(path)))
                .ToList();

            using var pkg = new ExcelPackage();

            // One chart sheet per chart definition
            foreach (var chart in Charts)
                AddChartSheet(pkg, chart, servers);

            pkg.SaveAs(new FileInfo(outputXlsxPath));
        }

        // ── Chart sheet builder ───────────────────────────────────────────────

        private static void AddChartSheet(
            ExcelPackage pkg, ChartDef def, List<ServerData> servers)
        {
            // Safe sheet name (Excel limit: 31 chars, no special chars)
            string sheetName = Sanitise(def.Title, 31);
            var ws = pkg.Workbook.Worksheets.Add(sheetName);

            // ── Collect data first ────────────────────────────────────────────
            List<string>? timestamps = null;
            var seriesValues = new List<(string Label, List<double> Values)>();

            foreach (var srv in servers)
            {
                var (ts, vals) = ExtractCounter(srv.Rows, def);
                if (timestamps == null && ts.Count > 0)
                    timestamps = ts;
                seriesValues.Add((srv.Label, vals));
            }

            if (timestamps == null || timestamps.Count == 0)
            {
                ws.Cells[1, 1].Value = $"No data found for: {def.SearchKeyword}";
                return;
            }

            int dataCols = timestamps.Count;

            // ── Write data table (transposed) ─────────────────────────────────
            // Col A        = row label ("Timestamp", server names)
            // Col B…N      = one column per time sample
            //
            // Row 1: "Timestamp"  | t1 | t2 | t3 | …
            // Row 2: "Server1"    | v  | v  | v  | …
            // Row 3: "Server2"    | v  | v  | v  | …

            // Row 1 — header label
            ws.Cells[1, 1].Value = "Timestamp";
            ws.Cells[1, 1].Style.Font.Bold = true;
            ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x1E, 0x40, 0xAF));
            ws.Cells[1, 1].Style.Font.Color.SetColor(Color.White);

            // Row 1 — timestamps across columns
            for (int c = 0; c < timestamps.Count; c++)
            {
                var cell = ws.Cells[1, c + 2];
                if (DateTime.TryParse(timestamps[c], CultureInfo.InvariantCulture,
                        DateTimeStyles.None, out var dt))
                {
                    cell.Value = dt;
                    cell.Style.Numberformat.Format = "hh:mm:ss";
                }
                else
                {
                    cell.Value = timestamps[c];
                }
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x1E, 0x40, 0xAF));
                cell.Style.Font.Color.SetColor(Color.White);
                cell.Style.Font.Bold = true;
            }

            // Rows 2…N — one row per server
            for (int s = 0; s < seriesValues.Count; s++)
            {
                int row = s + 2;
                var (label, vals) = seriesValues[s];

                // Col A = server label
                ws.Cells[row, 1].Value = label;
                ws.Cells[row, 1].Style.Font.Bold = true;
                ws.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(
                    row % 2 == 0
                        ? Color.FromArgb(0xDB, 0xEA, 0xFE)   // light blue for even rows
                        : Color.FromArgb(0xF3, 0xF4, 0xF6));  // light grey for odd
                ws.Cells[row, 1].Style.Font.Color.SetColor(Color.FromArgb(0x1E, 0x3A, 0x8A));

                // Col B… = values
                for (int c = 0; c < vals.Count && c < dataCols; c++)
                {
                    var cell = ws.Cells[row, c + 2];
                    cell.Value = vals[c];
                    cell.Style.Numberformat.Format = "0.00";
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(
                        row % 2 == 0
                            ? Color.FromArgb(0xEF, 0xF6, 0xFF)
                            : Color.FromArgb(0xF9, 0xFA, 0xFB));
                }
            }

            ws.Column(1).Width = Math.Max(14, seriesValues.Max(s => s.Label.Length) + 2);
            for (int c = 2; c <= dataCols + 1; c++)
                ws.Column(c).Width = 10;

            // ── Add line chart immediately below the data rows ────────────────
            int chartAnchorRow = seriesValues.Count + 3;   // e.g. 3 servers → row 6

            var chart = ws.Drawings.AddChart(def.Title, eChartType.Line) as ExcelLineChart;
            if (chart == null) return;

            chart.Title.Text = def.Title;
            chart.SetPosition(chartAnchorRow, 0, 0, 0);
            chart.SetSize(900, 380);

            // Colour palette for multiple series lines
            var lineColors = new[]
            {
                Color.FromArgb(0x25, 0x63, 0xEB),   // blue
                Color.FromArgb(0xDC, 0x26, 0x26),   // red
                Color.FromArgb(0x16, 0xA3, 0x4A),   // green
                Color.FromArgb(0xD9, 0x77, 0x06),   // amber
                Color.FromArgb(0x70, 0x3A, 0xED),   // purple
                Color.FromArgb(0x06, 0x96, 0x88),   // teal
            };

            // One series per server — data is in rows, X axis from row 1 cols 2…
            for (int s = 0; s < seriesValues.Count; s++)
            {
                int row = s + 2;
                var ser = (ExcelLineChartSerie)chart.Series.Add(
                    ws.Cells[row, 2, row, dataCols + 1],    // Y values (entire data row)
                    ws.Cells[1, 2, 1, dataCols + 1]);   // X (timestamp row)
                ser.Header = seriesValues[s].Label;

                // Set line colour and width (1.5pt)
                var color = lineColors[s % lineColors.Length];
                ser.Border.Fill.Color = color;
                ser.Border.Width = 1.5;

                // Hide the dot markers so only the line shows
                ser.Marker.Style = eMarkerStyle.None;
            }

            chart.Legend.Position = eLegendPosition.Bottom;
            chart.XAxis.Title.Text = "Time";
            chart.YAxis.Title.Text = def.ConvertToMbps ? "Mbps"
                                   : def.InvertIdle ? "%"
                                   : def.Title.Contains("MB") ? "MB"
                                   : def.Title.Contains("sec") ? "/ sec"
                                   : "";
        }

        // ── Counter extraction ────────────────────────────────────────────────

        private static (List<string> Timestamps, List<double> Values)
            ExtractCounter(List<string[]> rows, ChartDef def)
        {
            if (rows.Count < 2)
                return (new(), new());

            string[] headers = rows[0];

            // Find the column index whose header contains the search keyword
            int colIdx = -1;
            for (int c = 0; c < headers.Length; c++)
            {
                if (headers[c].IndexOf(def.SearchKeyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    colIdx = c;
                    break;
                }
            }

            if (colIdx < 0)
                return (new(), new());

            // relog row 1 (index 1) is a units row — skip it; data starts at row 2 (index 2)
            int dataStart = rows.Count > 1
                ? (IsUnitsRow(rows[1]) ? 2 : 1)
                : 1;

            var timestamps = new List<string>();
            var values = new List<double>();

            for (int r = dataStart; r < rows.Count; r++)
            {
                string[] row = rows[r];
                if (row.Length == 0 || string.IsNullOrWhiteSpace(row[0])) continue;

                timestamps.Add(row[0].Trim().Trim('"'));

                double val = 0;
                if (colIdx < row.Length)
                    double.TryParse(row[colIdx].Trim().Trim('"'),
                        NumberStyles.Any, CultureInfo.InvariantCulture, out val);

                // Transformations
                if (def.ConvertToMbps) val = val / 1_000_000.0 * 8.0;   // bytes/s → Mbps
                if (def.InvertIdle) val = 100.0 - val;                // % Idle → % Busy

                values.Add(Math.Round(val, 3));
            }

            return (timestamps, values);
        }

        // ── CSV parsing ───────────────────────────────────────────────────────

        private static List<string[]> ParseCsv(string path)
        {
            var result = new List<string[]>();
            foreach (var line in File.ReadAllLines(path))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                result.Add(SplitCsvLine(line));
            }
            return result;
        }

        private static string[] SplitCsvLine(string line)
        {
            var fields = new List<string>();
            var sb = new System.Text.StringBuilder();
            bool inQ = false;
            foreach (char c in line)
            {
                if (inQ)
                {
                    if (c == '"') inQ = false;
                    else sb.Append(c);
                }
                else
                {
                    if (c == '"') inQ = true;
                    else if (c == ',') { fields.Add(sb.ToString()); sb.Clear(); }
                    else sb.Append(c);
                }
            }
            fields.Add(sb.ToString());
            return fields.ToArray();
        }

        // relog row 1 contains unit strings like "(PDH-CSV 4.0)" or empty strings
        private static bool IsUnitsRow(string[] row) =>
            row.Length > 0 && (row[0].Contains("PDH") || row[0].Contains("(") ||
            (row.Length > 1 && string.IsNullOrWhiteSpace(row[1])));

        private static string Sanitise(string name, int maxLen)
        {
            var sb = new System.Text.StringBuilder();
            foreach (char c in name)
                sb.Append("[]:\\*?/".Contains(c) ? '_' : c);
            string s = sb.ToString().Trim();
            return s.Length > maxLen ? s[..maxLen] : s;
        }

        // ── Internal model ────────────────────────────────────────────────────

        private record ServerData(string Label, List<string[]> Rows);
    }
}
