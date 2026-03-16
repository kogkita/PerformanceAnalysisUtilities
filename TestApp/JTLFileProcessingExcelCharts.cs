using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Text;

namespace TestApp
{
    public static class JTLFileProcessingExcelCharts
    {
        // Chart dimensions
        private const long ChartW = 1400L * 9525L;  // 1400 px wide
        private const long ScaleChartH = 55L * 9525L;  // row 2: scale bar (taller for axis labels)
        private const long MiniChartH = 55L * 9525L;  // rows 3+: bar charts
        private const double TitleRowHt = 20.0;           // row 1 height (pt)
        private const double ScaleRowHt = 42.0;           // row 2 height (pt)
        private const double MiniRowHt = 42.0;           // rows 3+ height (pt)

        // ── Public API ────────────────────────────────────────────────────────

        public static void AddMiniChartsAndSave(
            ExcelPackage package,
            ExcelWorksheet dataSheet,
            List<JTLFileProcessingRecord> records,
            string xlsxPath)
        {
            // records is already sorted by avg desc (passed from Convert)
            int n = records.Count;

            // ── Build JTL Charts sheet ────────────────────────────────────────
            var cs = package.Workbook.Worksheets.Add("JTL Charts");
            cs.Column(1).Width = 42;   // Col A: transaction names

            // Row 1 — title
            cs.Row(1).Height = TitleRowHt;
            cs.Cells[1, 1].Value = "Transaction Latency \u2013 Average vs 90th Percentile (Seconds)  |  Scale: 0 \u2013 60 s  (values >60 s shown capped at 65 s with actual label)";
            cs.Cells[1, 1].Style.Font.Bold = true;
            cs.Cells[1, 1].Style.Font.Size = 12;
            cs.Cells[1, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // Row 2 — scale row: col A has coloured rich-text legend key
            cs.Row(2).Height = ScaleRowHt;
            cs.Cells[2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            // EPPlus RichText: "Scale   " normal, "■ Avg" blue, "    ■ P90" orange
            var rt = cs.Cells[2, 1].RichText;
            var rtScale = rt.Add("Scale    ");
            rtScale.Bold = true;
            rtScale.Color = System.Drawing.Color.Black;
            var rtAvgSq = rt.Add("\u25A0 ");
            rtAvgSq.Bold = true;
            rtAvgSq.Color = System.Drawing.Color.FromArgb(0x20, 0x6B, 0xA3); // Excel default series 1 blue
            var rtAvgLbl = rt.Add("Avg");
            rtAvgLbl.Bold = true;
            rtAvgLbl.Color = System.Drawing.Color.Black;
            var rtP90Sep = rt.Add("    ");
            rtP90Sep.Color = System.Drawing.Color.Black;
            var rtP90Sq = rt.Add("\u25A0 ");
            rtP90Sq.Bold = true;
            rtP90Sq.Color = System.Drawing.Color.FromArgb(0xE3, 0x6C, 0x09); // Excel default series 2 orange
            var rtP90Lbl = rt.Add("P90");
            rtP90Lbl.Bold = true;
            rtP90Lbl.Color = System.Drawing.Color.Black;

            // Register the scale-axis chart shell (index 0 = chart1)
            var scaleChart = (OfficeOpenXml.Drawing.Chart.ExcelBarChart)
                cs.Drawings.AddChart("ScaleAxis",
                    OfficeOpenXml.Drawing.Chart.eChartType.BarClustered);
            scaleChart.SetPosition(1, 0, 1, 0);   // 0-based row 1 = sheet row 2, col B = 1
            scaleChart.SetSize(1, 1);

            // Rows 3+ — one row per transaction, in JTL Results sheet order
            for (int i = 0; i < n; i++)
            {
                int row = 3 + i;
                cs.Row(row).Height = MiniRowHt;
                cs.Cells[row, 1].Value = records[i].TransactionName;
                cs.Cells[row, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                // Register EPPlus chart shell (indices 1..n = chart2..chartN+1)
                var c = (OfficeOpenXml.Drawing.Chart.ExcelBarChart)
                    cs.Drawings.AddChart($"C{i}",
                        OfficeOpenXml.Drawing.Chart.eChartType.BarClustered);
                c.SetPosition(row - 1, 0, 1, 0);  // 0-based row, col B
                c.SetSize(1, 1);
            }

            // Freeze rows 1-2 so title + scale stay visible while scrolling
            cs.View.FreezePanes(3, 1);   // freeze above row 3

            package.SaveAs(new FileInfo(xlsxPath));

            // Inject correct chart XML + drawing XML into the saved ZIP
            InjectAllCharts(xlsxPath, records);
        }

        public static void InjectChartForSheet(
            string xlsxPath,
            string sheetName,
            List<JTLFileProcessingRecord> records)
        {
            InjectAllCharts(xlsxPath, records);
        }

        // ── ZIP injection ─────────────────────────────────────────────────────

        private static void InjectAllCharts(string xlsxPath, List<JTLFileProcessingRecord> records)
        {
            int n = records.Count;

            using var pkg = Package.Open(xlsxPath, FileMode.Open, FileAccess.ReadWrite);

            // Collect and sort chart parts (chart1.xml, chart2.xml, ...)
            var chartParts = new List<PackagePart>();
            PackagePart? drawingPart = null;

            foreach (var part in pkg.GetParts())
            {
                var u = part.Uri.ToString();
                if (u.StartsWith("/xl/charts/chart", System.StringComparison.OrdinalIgnoreCase)
                    && u.EndsWith(".xml", System.StringComparison.OrdinalIgnoreCase))
                    chartParts.Add(part);
                if (u.StartsWith("/xl/drawings/drawing", System.StringComparison.OrdinalIgnoreCase)
                    && u.EndsWith(".xml", System.StringComparison.OrdinalIgnoreCase)
                    && !u.Contains("_rels"))
                    drawingPart = part;
            }

            chartParts.Sort((a, b) =>
                ExtractNum(a.Uri.ToString()).CompareTo(ExtractNum(b.Uri.ToString())));

            // Read rIds from drawing
            var rIds = new List<string>();
            if (drawingPart != null)
            {
                string d;
                using (var sr = new StreamReader(
                    drawingPart.GetStream(FileMode.Open, FileAccess.Read)))
                    d = sr.ReadToEnd();
                foreach (System.Text.RegularExpressions.Match m in
                    System.Text.RegularExpressions.Regex.Matches(d, @"r:id=""([^""]+)"""))
                    rIds.Add(m.Groups[1].Value);
            }

            // chart index 0 = scale axis chart (row 2)
            if (chartParts.Count > 0)
            {
                var xml = BuildScaleChartXml();
                var bytes = Encoding.UTF8.GetBytes(xml);
                using var s = chartParts[0].GetStream(FileMode.Create, FileAccess.Write);
                s.Write(bytes, 0, bytes.Length);
            }

            // chart indices 1..n = transaction mini charts (rows 3+)
            for (int i = 0; i < n && (i + 1) < chartParts.Count; i++)
            {
                var xml = BuildMiniChartXml(records[i], i + 1);
                var bytes = Encoding.UTF8.GetBytes(xml);
                using var s = chartParts[i + 1].GetStream(FileMode.Create, FileAccess.Write);
                s.Write(bytes, 0, bytes.Length);
            }

            // Overwrite drawing XML with correct anchors
            if (drawingPart != null)
            {
                var dxml = BuildDrawingXml(n, rIds);
                var bytes = Encoding.UTF8.GetBytes(dxml);
                using var s = drawingPart.GetStream(FileMode.Create, FileAccess.Write);
                s.Write(bytes, 0, bytes.Length);
            }
        }

        // ── Scale axis chart (row 2) ──────────────────────────────────────────
        // A bar chart with no data series — just renders the x-axis 0-60s.

        private static string BuildScaleChartXml()
        {
            var sb = new StringBuilder(1024);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"");
            sb.Append(" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"");
            sb.Append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sb.Append("<c:lang val=\"en-US\"/><c:roundedCorners val=\"0\"/>");
            sb.Append("<c:chart><c:autoTitleDeleted val=\"1\"/>");
            sb.Append("<c:plotArea><c:layout/>");
            sb.Append("<c:barChart><c:barDir val=\"bar\"/><c:grouping val=\"clustered\"/>");
            sb.Append("<c:varyColors val=\"0\"/>");
            // Two dummy invisible series — force axis to render, no visible bars
            sb.Append("<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>");
            sb.Append("<c:tx><c:v>Avg</c:v></c:tx>");
            sb.Append("<c:invertIfNegative val=\"0\"/>");
            sb.Append("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>");
            sb.Append("<c:cat><c:strLit><c:ptCount val=\"1\"/>");
            sb.Append("<c:pt idx=\"0\"><c:v> </c:v></c:pt></c:strLit></c:cat>");
            sb.Append("<c:val><c:numLit><c:ptCount val=\"1\"/>");
            sb.Append("<c:pt idx=\"0\"><c:v>0</c:v></c:pt></c:numLit></c:val></c:ser>");
            sb.Append("<c:ser><c:idx val=\"1\"/><c:order val=\"1\"/>");
            sb.Append("<c:tx><c:v>P90</c:v></c:tx>");
            sb.Append("<c:invertIfNegative val=\"0\"/>");
            sb.Append("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>");
            sb.Append("<c:cat><c:strLit><c:ptCount val=\"1\"/>");
            sb.Append("<c:pt idx=\"0\"><c:v> </c:v></c:pt></c:strLit></c:cat>");
            sb.Append("<c:val><c:numLit><c:ptCount val=\"1\"/>");
            sb.Append("<c:pt idx=\"0\"><c:v>0</c:v></c:pt></c:numLit></c:val></c:ser>");
            sb.Append("<c:gapWidth val=\"50\"/>");
            sb.Append("<c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>");
            // catAx hidden
            sb.Append("<c:catAx><c:axId val=\"1\"/>");
            sb.Append("<c:scaling><c:orientation val=\"minMax\"/></c:scaling>");
            sb.Append("<c:delete val=\"1\"/><c:axPos val=\"l\"/>");
            sb.Append("<c:crossAx val=\"2\"/></c:catAx>");
            // valAx: 0-120s, 20s intervals, labels visible
            sb.Append("<c:valAx><c:axId val=\"2\"/>");
            sb.Append("<c:scaling><c:orientation val=\"minMax\"/>");
            sb.Append("<c:min val=\"0\"/><c:max val=\"70\"/></c:scaling>");
            sb.Append("<c:delete val=\"0\"/><c:axPos val=\"b\"/>");
            sb.Append("<c:majorGridlines/>");
            sb.Append("<c:numFmt formatCode=\"[&lt;70]0;[=70]&quot;&quot;;0\" sourceLinked=\"0\"/>");
            sb.Append("<c:tickLblPos val=\"low\"/>");
            sb.Append("<c:crossAx val=\"1\"/><c:crosses val=\"min\"/>");
            sb.Append("<c:crossBetween val=\"between\"/>");
            sb.Append("<c:majorUnit val=\"10\"/></c:valAx>");
            sb.Append("</c:plotArea>");
            // NO legend — colour key is in col A cell text instead, so plot area
            // width matches transaction charts exactly (no legend offset)
            sb.Append("<c:legend><c:delete val=\"1\"/></c:legend>");
            sb.Append("<c:plotVisOnly val=\"1\"/><c:dispBlanksAs val=\"zero\"/>");
            sb.Append("</c:chart>");
            sb.Append("<c:printSettings><c:headerFooter/>");
            sb.Append("<c:pageMargins b=\"0.25\" l=\"0.25\" r=\"0.25\" t=\"0.25\" header=\"0.3\" footer=\"0.3\"/>");
            sb.Append("<c:pageSetup/></c:printSettings></c:chartSpace>");
            return sb.ToString();
        }

        // ── Mini bar chart (rows 3+) — bars only, NO axis labels ─────────────

        private static string BuildMiniChartXml(JTLFileProcessingRecord r, int idx)
        {
            const double MaxScale = 70.0;  // axis max
            const double CapAt = 65.0;  // bar capped here for overflow values

            double avgReal = System.Math.Round(r.Average / 1000.0, 3);
            double p90Real = System.Math.Round(r.P90 / 1000.0, 3);
            // Bar values: cap at 65 so they never exceed the visible plot area
            double avgBar = System.Math.Min(avgReal, CapAt);
            double p90Bar = System.Math.Min(p90Real, CapAt);

            string avgBarStr = avgBar.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string p90BarStr = p90Bar.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string avgLblStr = avgReal.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string p90LblStr = p90Real.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string a1 = (idx * 2 + 1).ToString();
            string a2 = (idx * 2 + 2).ToString();

            var sb = new StringBuilder(1200);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"");
            sb.Append(" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"");
            sb.Append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sb.Append("<c:lang val=\"en-US\"/><c:roundedCorners val=\"0\"/>");
            sb.Append("<c:chart><c:autoTitleDeleted val=\"1\"/>");
            sb.Append("<c:plotArea><c:layout/>");
            sb.Append("<c:barChart><c:barDir val=\"bar\"/><c:grouping val=\"clustered\"/>");
            sb.Append("<c:varyColors val=\"0\"/>");

            // Average series — bar capped at 65, label shows real value
            sb.Append($"<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>");
            sb.Append("<c:tx><c:v>Avg</c:v></c:tx>");
            sb.Append("<c:invertIfNegative val=\"0\"/>");
            sb.Append("<c:dLbls>");
            // Custom label for the single point showing REAL value
            sb.Append("<c:dLbl><c:idx val=\"0\"/>");
            sb.Append("<c:tx><c:rich><a:bodyPr/><a:lstStyle/>");
            sb.Append($"<a:p><a:r><a:t>{avgLblStr}</a:t></a:r></a:p></c:rich></c:tx>");
            sb.Append("<c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/>");
            sb.Append("<c:showCatName val=\"0\"/><c:showSerName val=\"0\"/>");
            sb.Append("<c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/></c:dLbl>");
            sb.Append("<c:dLblPos val=\"outEnd\"/>");
            sb.Append("<c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/>");
            sb.Append("<c:showCatName val=\"0\"/><c:showSerName val=\"0\"/>");
            sb.Append("<c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/></c:dLbls>");
            sb.Append("<c:cat><c:strLit><c:ptCount val=\"1\"/>");
            sb.Append("<c:pt idx=\"0\"><c:v> </c:v></c:pt></c:strLit></c:cat>");
            sb.Append($"<c:val><c:numLit><c:ptCount val=\"1\"/>");
            sb.Append($"<c:pt idx=\"0\"><c:v>{avgBarStr}</c:v></c:pt></c:numLit></c:val></c:ser>");

            // P90 series — bar capped at 65, label shows real value
            sb.Append($"<c:ser><c:idx val=\"1\"/><c:order val=\"1\"/>");
            sb.Append("<c:tx><c:v>P90</c:v></c:tx>");
            sb.Append("<c:invertIfNegative val=\"0\"/>");
            sb.Append("<c:dLbls>");
            sb.Append("<c:dLbl><c:idx val=\"0\"/>");
            sb.Append("<c:tx><c:rich><a:bodyPr/><a:lstStyle/>");
            sb.Append($"<a:p><a:r><a:t>{p90LblStr}</a:t></a:r></a:p></c:rich></c:tx>");
            sb.Append("<c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/>");
            sb.Append("<c:showCatName val=\"0\"/><c:showSerName val=\"0\"/>");
            sb.Append("<c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/></c:dLbl>");
            sb.Append("<c:dLblPos val=\"outEnd\"/>");
            sb.Append("<c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/>");
            sb.Append("<c:showCatName val=\"0\"/><c:showSerName val=\"0\"/>");
            sb.Append("<c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/></c:dLbls>");
            sb.Append("<c:cat><c:strLit><c:ptCount val=\"1\"/>");
            sb.Append("<c:pt idx=\"0\"><c:v> </c:v></c:pt></c:strLit></c:cat>");
            sb.Append($"<c:val><c:numLit><c:ptCount val=\"1\"/>");
            sb.Append($"<c:pt idx=\"0\"><c:v>{p90BarStr}</c:v></c:pt></c:numLit></c:val></c:ser>");

            sb.Append("<c:gapWidth val=\"50\"/>");
            sb.Append($"<c:axId val=\"{a1}\"/><c:axId val=\"{a2}\"/></c:barChart>");

            // catAx — hidden
            sb.Append($"<c:catAx><c:axId val=\"{a1}\"/>");
            sb.Append("<c:scaling><c:orientation val=\"minMax\"/></c:scaling>");
            sb.Append("<c:delete val=\"1\"/><c:axPos val=\"l\"/>");
            sb.Append($"<c:crossAx val=\"{a2}\"/></c:catAx>");

            // valAx — same 0-120s range to keep bars proportional, but completely hidden
            sb.Append($"<c:valAx><c:axId val=\"{a2}\"/>");
            sb.Append("<c:scaling><c:orientation val=\"minMax\"/>");
            sb.Append("<c:min val=\"0\"/><c:max val=\"70\"/></c:scaling>");
            sb.Append("<c:delete val=\"0\"/><c:axPos val=\"b\"/>");
            sb.Append("<c:tickLblPos val=\"none\"/>");
            sb.Append("<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>"); // hide axis line
            sb.Append($"<c:crossAx val=\"{a1}\"/><c:crosses val=\"min\"/>");
            sb.Append("<c:crossBetween val=\"between\"/>");
            sb.Append("<c:majorUnit val=\"10\"/></c:valAx>");

            sb.Append("</c:plotArea>");
            sb.Append("<c:legend><c:delete val=\"1\"/></c:legend>"); // no legend on transaction charts
            sb.Append("<c:plotVisOnly val=\"1\"/><c:dispBlanksAs val=\"zero\"/>");
            sb.Append("</c:chart>");
            sb.Append("<c:printSettings><c:headerFooter/>");
            sb.Append("<c:pageMargins b=\"0.25\" l=\"0.25\" r=\"0.25\" t=\"0.25\" header=\"0.3\" footer=\"0.3\"/>");
            sb.Append("<c:pageSetup/></c:printSettings></c:chartSpace>");
            return sb.ToString();
        }

        // ── Drawing XML ───────────────────────────────────────────────────────

        private static string BuildDrawingXml(int n, List<string> rIds)
        {
            const string xdrNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
            const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
            const string cNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";

            var sb = new StringBuilder((n + 1) * 500);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append($"<xdr:wsDr xmlns:xdr=\"{xdrNs}\" xmlns:a=\"{aNs}\"");
            sb.Append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");

            // Total charts = 1 (scale) + n (transactions)
            int total = n + 1;
            for (int i = 0; i < total; i++)
            {
                string rId = i < rIds.Count ? rIds[i] : $"rId{i + 1}";
                long cy = i == 0 ? ScaleChartH : MiniChartH;
                int row = i + 1;    // 0-based: row1=scale(sheet row2), row2=tx0(sheet row3)...
                string name = i == 0 ? "ScaleAxis" : $"C{i - 1}";

                sb.Append("<xdr:oneCellAnchor>");
                sb.Append($"<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>");
                sb.Append($"<xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>");
                sb.Append($"<xdr:ext cx=\"{ChartW}\" cy=\"{cy}\"/>");
                sb.Append("<xdr:graphicFrame macro=\"\">");
                sb.Append("<xdr:nvGraphicFramePr>");
                sb.Append($"<xdr:cNvPr id=\"{i + 2}\" name=\"{name}\"/>");
                sb.Append("<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>");
                sb.Append($"<xdr:xfrm><a:off x=\"0\" y=\"0\"/>");
                sb.Append($"<a:ext cx=\"{ChartW}\" cy=\"{cy}\"/></xdr:xfrm>");
                sb.Append("<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">");
                sb.Append($"<c:chart xmlns:c=\"{cNs}\" r:id=\"{rId}\"/>");
                sb.Append("</a:graphicData></a:graphic></xdr:graphicFrame>");
                sb.Append("<xdr:clientData/></xdr:oneCellAnchor>");
            }

            sb.Append("</xdr:wsDr>");
            return sb.ToString();
        }

        private static int ExtractNum(string uri)
        {
            var m = System.Text.RegularExpressions.Regex.Match(uri, @"chart(\d+)\.xml");
            return m.Success ? int.Parse(m.Groups[1].Value) : 0;
        }

        private static string EscapeXml(string s) =>
            s.Replace("&", "&amp;").Replace("<", "&lt;")
             .Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
    }
}
