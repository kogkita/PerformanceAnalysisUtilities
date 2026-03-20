using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;

namespace TestApp
{
    /// <summary>
    /// Produces an Excel workbook from one or more parsed NmonFile objects.
    /// Each nmon section (CPU_ALL, MEM, DISKBUSY, NET, PAGE, PROC, TOP, AAA, BBB)
    /// gets its own sheet with data + a line chart.
    /// </summary>
    public static class NmonExcelProducer
    {
        // ── Colours ───────────────────────────────────────────────────────────
        private static readonly Color HdrBg   = Color.FromArgb(0x1E, 0x40, 0xAF);
        private static readonly Color HdrFg   = Color.White;
        private static readonly Color EvenRow = Color.FromArgb(0xF3, 0xF4, 0xF6);
        private static readonly Color OddRow  = Color.White;
        private static readonly Color SummBg  = Color.FromArgb(0xDB, 0xEA, 0xFE);
        private static readonly Color SummFg  = Color.FromArgb(0x1E, 0x3A, 0x8A);

        private static readonly Color[] LineColors =
        {
            Color.FromArgb(0x25, 0x63, 0xEB), Color.FromArgb(0xDC, 0x26, 0x26),
            Color.FromArgb(0x16, 0xA3, 0x4A), Color.FromArgb(0xD9, 0x77, 0x06),
            Color.FromArgb(0x70, 0x3A, 0xED), Color.FromArgb(0x06, 0x96, 0x88),
            Color.FromArgb(0xDB, 0x27, 0x77), Color.FromArgb(0xEA, 0x58, 0x0C),
        };

        // ── Sheet definitions ─────────────────────────────────────────────────

        private record SheetDef(
            string Tag,
            string SheetName,
            string ChartTitle,
            string YAxisLabel,
            string[] SeriesColumns,   // which columns to chart (empty = all numeric)
            bool    HasChart);

        private static readonly SheetDef[] SheetDefs =
        {
            new("CPU_ALL",   "CPU_ALL",   "CPU Utilisation (%)",       "%",      new[]{"User%","Sys%","Wait%"}, true),
            new("MEM",       "MEM",       "Memory (MB)",               "MB",     new[]{"memfree","memtotal"},   true),
            new("DISKBUSY",  "DISKBUSY",  "Disk Busy (%)",             "%",      Array.Empty<string>(),         true),
            new("DISKREAD",  "DISKREAD",  "Disk Read (KB/s)",          "KB/s",   Array.Empty<string>(),         true),
            new("DISKWRITE", "DISKWRITE", "Disk Write (KB/s)",         "KB/s",   Array.Empty<string>(),         true),
            new("DISKXFER",  "DISKXFER",  "Disk Transfers/s",          "xfer/s", Array.Empty<string>(),         true),
            new("NET",       "NET",       "Network (KB/s)",            "KB/s",   Array.Empty<string>(),         true),
            new("NETPACKET", "NETPACKET", "Network Packets/s",         "pkt/s",  Array.Empty<string>(),         true),
            new("PAGE",      "PAGE",      "Paging (pages/s)",          "pg/s",   new[]{"pgin","pgout"},         true),
            new("PROC",      "PROC",      "Run Queue & Process Stats", "count",  new[]{"Runnable","Blocked"},   true),
            new("TOP",       "TOP",       "Top Processes (CPU%)",      "%",      Array.Empty<string>(),         false),
            new("LPAR",      "LPAR",      "LPAR Stats",                "",       Array.Empty<string>(),         true),
            new("VM",        "VM",        "Virtual Memory",            "",       Array.Empty<string>(),         true),
            new("JFSFILE",   "JFSFILE",   "Filesystem Usage (%)",      "%",      Array.Empty<string>(),         true),
        };

        // ── Public entry point ────────────────────────────────────────────────

        /// <summary>
        /// Parse all nmon files and write a combined Excel workbook.
        /// If multiple files, a summary sheet is added first.
        /// </summary>
        public static void Produce(
            IList<string> nmonPaths,
            string outputXlsxPath,
            IProgress<string>? progress = null)
        {
            ExcelPackage.License.SetNonCommercialPersonal("nmon Analyser");

            // Parse all files
            var nmonFiles = new List<NmonFile>();
            foreach (var path in nmonPaths)
            {
                progress?.Report($"Parsing {Path.GetFileName(path)}…");
                nmonFiles.Add(NmonParser.Parse(path));
            }

            using var pkg = new ExcelPackage();

            if (nmonFiles.Count > 1)
            {
                progress?.Report("Writing summary sheet…");
                WriteSummarySheet(pkg, nmonFiles);
            }

            // Write each file as its own group of sheets
            foreach (var nf in nmonFiles)
            {
                string prefix = nmonFiles.Count > 1
                    ? SanitiseName(nf.Meta.Host.Length > 0 ? nf.Meta.Host : Path.GetFileNameWithoutExtension(nf.Meta.FileName), 10) + " – "
                    : "";

                progress?.Report($"Writing sheets for {nf.Meta.Host}…");
                WriteAaaSheet(pkg, nf, prefix);
                WriteBbbSheet(pkg, nf, prefix);

                foreach (var def in SheetDefs)
                {
                    if (!nf.Sections.TryGetValue(def.Tag, out var section)) continue;
                    if (section.Rows.Count == 0) continue;

                    progress?.Report($"  {prefix}{def.SheetName}…");
                    WriteDataSheet(pkg, section, def, prefix);
                }
            }

            progress?.Report("Saving workbook…");
            pkg.SaveAs(new FileInfo(outputXlsxPath));
        }

        // ── Summary sheet (multi-file) ────────────────────────────────────────

        private static void WriteSummarySheet(ExcelPackage pkg, List<NmonFile> files)
        {
            var ws = pkg.Workbook.Worksheets.Add("Summary");

            ws.Cells[1, 1].Value = "nmon Analysis Summary";
            ws.Cells[1, 1].Style.Font.Size = 16;
            ws.Cells[1, 1].Style.Font.Bold = true;
            ws.Cells[1, 1].Style.Font.Color.SetColor(HdrBg);

            ws.Cells[2, 1].Value = $"Generated: {DateTime.Now:yyyy-MM-dd HH:mm}";
            ws.Cells[2, 1].Style.Font.Color.SetColor(Color.Gray);

            // Table header
            int row = 4;
            WriteRow(ws, row, new[] { "Host", "File", "Date", "Time", "OS", "Version", "Snapshots", "Interval (s)" }, isHeader: true);
            row++;

            foreach (var nf in files)
            {
                var m = nf.Meta;
                ws.Cells[row, 1].Value = m.Host;
                ws.Cells[row, 2].Value = m.FileName;
                ws.Cells[row, 3].Value = m.Date;
                ws.Cells[row, 4].Value = m.Time;
                ws.Cells[row, 5].Value = m.OS;
                ws.Cells[row, 6].Value = m.Version;
                ws.Cells[row, 7].Value = m.Snapshots;
                ws.Cells[row, 8].Value = m.Interval;
                if (row % 2 == 0)
                    ws.Cells[row, 1, row, 8].Style.Fill.SetBackground(EvenRow);
                row++;
            }

            ws.Cells.AutoFitColumns(10, 60);
        }

        // ── AAA sheet ─────────────────────────────────────────────────────────

        private static void WriteAaaSheet(ExcelPackage pkg, NmonFile nf, string prefix)
        {
            string sheetName = Unique(pkg, prefix + "AAA");
            var ws = pkg.Workbook.Worksheets.Add(sheetName);

            WriteRow(ws, 1, new[] { "Key", "Value" }, isHeader: true);

            var m = nf.Meta;
            var items = new[]
            {
                ("Host",        m.Host),
                ("File",        m.FileName),
                ("Date",        m.Date),
                ("Time",        m.Time),
                ("OS",          m.OS),
                ("Version",     m.Version),
                ("Snapshots",   m.Snapshots.ToString()),
                ("Interval",    m.Interval > 0 ? $"{m.Interval}s" : ""),
            };

            int row = 2;
            foreach (var (k, v) in items)
            {
                ws.Cells[row, 1].Value = k;
                ws.Cells[row, 2].Value = v;
                if (row % 2 == 0)
                    ws.Cells[row, 1, row, 2].Style.Fill.SetBackground(EvenRow);
                row++;
            }

            // Also add AAA section raw data if it exists
            if (nf.Sections.TryGetValue("AAA", out var aaaSec))
            {
                row++;
                foreach (var dataRow in aaaSec.Rows)
                {
                    for (int c = 0; c < dataRow.Length; c++)
                        ws.Cells[row, c + 1].Value = dataRow[c];
                    row++;
                }
            }

            ws.Column(1).Width = 20;
            ws.Column(2).Width = 50;
        }

        // ── BBB sheet ─────────────────────────────────────────────────────────

        private static void WriteBbbSheet(ExcelPackage pkg, NmonFile nf, string prefix)
        {
            if (nf.Meta.BbbLines.Count == 0) return;
            string sheetName = Unique(pkg, prefix + "BBB");
            var ws = pkg.Workbook.Worksheets.Add(sheetName);

            ws.Cells[1, 1].Value = "System Configuration (BBB)";
            ws.Cells[1, 1].Style.Font.Bold = true;
            ws.Cells[1, 1].Style.Font.Color.SetColor(HdrBg);

            int row = 2;
            foreach (var line in nf.Meta.BbbLines)
            {
                ws.Cells[row, 1].Value = line;
                ws.Cells[row, 1].Style.Font.Size = 10;
                ws.Cells[row, 1].Style.Font.Name = "Consolas";
                row++;
            }
            ws.Column(1).Width = 120;
        }

        // ── Generic data sheet + chart ────────────────────────────────────────

        private static void WriteDataSheet(
            ExcelPackage pkg, NmonSection section, SheetDef def, string prefix)
        {
            string sheetName = Unique(pkg, prefix + def.SheetName);
            var ws = pkg.Workbook.Worksheets.Add(sheetName);

            int colCount  = section.Columns.Count;   // number of metrics (incl. Timestamp)
            int dataRows  = section.Rows.Count;       // number of time intervals

            // ── Transposed layout ──────────────────────────────────────────────
            // Row 1        : "Timestamp" | t1 | t2 | t3 | …
            // Row 2…N      : metric name | v  | v  | v  | …
            // Row N+2      : "Average"   | avg| avg| …
            // Row N+3      : "Maximum"   | max| max| …
            // Row N+5+     : chart (immediately visible)

            // ── Row 1: timestamp header + timestamps across columns ───────────
            StyleHeaderCell(ws.Cells[1, 1], "Timestamp");

            for (int r = 0; r < dataRows; r++)
            {
                var cell = ws.Cells[1, r + 2];
                var row  = section.Rows[r];
                if (DateTime.TryParse(row[0], CultureInfo.InvariantCulture,
                        DateTimeStyles.None, out var dt))
                {
                    cell.Value = dt;
                    cell.Style.Numberformat.Format = "hh:mm:ss";
                }
                else
                {
                    cell.Value = row[0];
                }
                StyleHeaderCell(cell);
            }

            // ── Rows 2…N: one row per metric ──────────────────────────────────
            for (int c = 1; c < colCount; c++)   // c=0 is Timestamp, skip it
            {
                int excelRow = c + 1;  // row 2 = first metric

                // Col A: metric name
                var nameCell = ws.Cells[excelRow, 1];
                nameCell.Value = section.Columns[c];
                nameCell.Style.Font.Bold = true;
                nameCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                nameCell.Style.Fill.BackgroundColor.SetColor(
                    excelRow % 2 == 0 ? Color.FromArgb(0xDB, 0xEA, 0xFE) : EvenRow);
                nameCell.Style.Font.Color.SetColor(SummFg);

                // Cols B…: values across time
                for (int r = 0; r < dataRows; r++)
                {
                    var row  = section.Rows[r];
                    var cell = ws.Cells[excelRow, r + 2];

                    // Alternating cell bg
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(
                        excelRow % 2 == 0 ? Color.FromArgb(0xEF, 0xF6, 0xFF) : OddRow);

                    if (c < row.Length &&
                        double.TryParse(row[c], NumberStyles.Any,
                            CultureInfo.InvariantCulture, out double d))
                    {
                        cell.Value = d;
                        cell.Style.Numberformat.Format = "0.00";
                    }
                    else if (c < row.Length)
                    {
                        cell.Value = row[c];
                    }
                }
            }

            // ── Avg / Max summary rows ────────────────────────────────────────
            if (dataRows > 0 && colCount > 1)
            {
                int avgRow  = colCount + 2;   // one blank gap row
                int maxRow  = colCount + 3;
                int lastCol = dataRows + 1;   // last data column (1-based)

                StyleSummaryLabel(ws.Cells[avgRow, 1], "Average");
                StyleSummaryLabel(ws.Cells[maxRow, 1], "Maximum");

                for (int r = 0; r < dataRows; r++)
                {
                    int col = r + 2;
                    string metricStart = $"{GetColLetter(col)}{2}";
                    string metricEnd   = $"{GetColLetter(col)}{colCount}";

                    ws.Cells[avgRow, col].Formula = $"=IFERROR(AVERAGE({metricStart}:{metricEnd}),\"\")";
                    ws.Cells[maxRow, col].Formula = $"=IFERROR(MAX({metricStart}:{metricEnd}),\"\")";
                    ws.Cells[avgRow, col].Style.Numberformat.Format = "0.00";
                    ws.Cells[maxRow, col].Style.Numberformat.Format = "0.00";
                }

                ws.Cells[avgRow, 1, avgRow, lastCol].Style.Fill.SetBackground(SummBg);
                ws.Cells[maxRow, 1, maxRow, lastCol].Style.Fill.SetBackground(SummBg);
                ws.Cells[avgRow, 1, maxRow, 1].Style.Font.Color.SetColor(SummFg);
                ws.Cells[avgRow, 1, maxRow, 1].Style.Font.Bold = true;
            }

            // Column A width = longest metric name
            int maxNameLen = section.Columns.Skip(1).Max(c => c.Length);
            ws.Column(1).Width = Math.Max(14, Math.Min(maxNameLen + 2, 30));
            for (int col = 2; col <= dataRows + 1; col++)
                ws.Column(col).Width = 10;

            // Freeze col A so metric names stay visible while scrolling right
            ws.View.FreezePanes(1, 2);

            // ── Chart immediately below data ───────────────────────────────────
            if (def.HasChart && dataRows > 1)
                AddLineChart(ws, section, def, dataRows, colCount);
        }

        // ── Line chart (transposed) ───────────────────────────────────────────

        private static void AddLineChart(
            ExcelWorksheet ws, NmonSection section, SheetDef def,
            int dataRows, int colCount)
        {
            // Determine which metric rows to plot (rows 2…N in the transposed sheet)
            var chartRows = new List<int>(); // 1-based excel row indices

            if (def.SeriesColumns.Length > 0)
            {
                foreach (var name in def.SeriesColumns)
                {
                    int idx = section.Columns.FindIndex(c =>
                        c.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0);
                    if (idx > 0) chartRows.Add(idx + 1); // row = col index + 1
                }
            }

            if (chartRows.Count == 0)
            {
                // Auto: all metric rows (rows 2 … min(colCount, 11))
                for (int row = 2; row <= Math.Min(colCount, 10); row++)
                    chartRows.Add(row);
            }

            if (chartRows.Count == 0) return;

            // Anchor chart just below data + summary rows — always visible
            int chartAnchorRow = colCount + 5;

            var chart = ws.Drawings.AddChart(def.ChartTitle, eChartType.Line) as ExcelLineChart;
            if (chart == null) return;

            chart.Title.Text = def.ChartTitle;
            chart.SetPosition(chartAnchorRow, 0, 0, 0);
            chart.SetSize(900, 360);

            // X axis: timestamp row (row 1, cols 2…dataRows+1)
            var xRange = ws.Cells[1, 2, 1, dataRows + 1];

            for (int i = 0; i < chartRows.Count; i++)
            {
                int row = chartRows[i];
                if (row > colCount) continue;
                var yRange = ws.Cells[row, 2, row, dataRows + 1];
                var ser    = (ExcelLineChartSerie)chart.Series.Add(yRange, xRange);
                ser.Header = section.Columns[row - 1];
                ser.Border.Fill.Color = LineColors[i % LineColors.Length];
                ser.Border.Width      = 1.5;
                ser.Marker.Style      = eMarkerStyle.None;
            }

            chart.Legend.Position  = eLegendPosition.Bottom;
            chart.XAxis.Title.Text = "Time";
            chart.YAxis.Title.Text = def.YAxisLabel;
        }

        // ── Cell style helpers ────────────────────────────────────────────────

        private static void StyleHeaderCell(ExcelRange cell, string? value = null)
        {
            if (value != null) cell.Value = value;
            cell.Style.Font.Bold = true;
            cell.Style.Font.Color.SetColor(HdrFg);
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(HdrBg);
        }

        private static void StyleSummaryLabel(ExcelRange cell, string value)
        {
            cell.Value = value;
            cell.Style.Font.Bold = true;
            cell.Style.Font.Color.SetColor(SummFg);
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(SummBg);
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static void WriteRow(ExcelWorksheet ws, int row, string[] values, bool isHeader)
        {
            for (int c = 0; c < values.Length; c++)
            {
                var cell = ws.Cells[row, c + 1];
                cell.Value = values[c];
                if (isHeader)
                {
                    cell.Style.Font.Bold = true;
                    cell.Style.Font.Color.SetColor(HdrFg);
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(HdrBg);
                }
            }
        }

        private static void SetBackground(this ExcelStyle style, Color color)
        {
            style.Fill.PatternType = ExcelFillStyle.Solid;
            style.Fill.BackgroundColor.SetColor(color);
        }

        private static string GetColLetter(int col)
        {
            string result = "";
            while (col > 0)
            {
                col--;
                result = (char)('A' + col % 26) + result;
                col /= 26;
            }
            return result;
        }

        private static string Unique(ExcelPackage pkg, string name)
        {
            name = SanitiseName(name, 31);
            string candidate = name;
            int n = 2;
            while (pkg.Workbook.Worksheets.Any(ws =>
                ws.Name.Equals(candidate, StringComparison.OrdinalIgnoreCase)))
            {
                string suffix = $" ({n++})";
                candidate = name[..Math.Min(name.Length, 31 - suffix.Length)] + suffix;
            }
            return candidate;
        }

        private static string SanitiseName(string name, int maxLen)
        {
            var sb = new System.Text.StringBuilder();
            foreach (char c in name)
                sb.Append("[]:\\*?/".Contains(c) ? '_' : c);
            string s = sb.ToString().Trim();
            return s.Length > maxLen ? s[..maxLen] : s;
        }
    }
}
