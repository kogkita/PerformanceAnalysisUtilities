using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Text;

namespace TestApp
{
    public static class ResponseTimeConverterExcelCharts
    {
        // Chart dimensions — same constants as JTL charts for visual consistency
        private const long ChartW = 1400L * 9525L;
        private const long ScaleChartH = 55L * 9525L;
        private const long MiniChartH = 55L * 9525L;
        private const double TitleRowHt = 20.0;
        private const double ScaleRowHt = 42.0;
        private const double MiniRowHt = 42.0;

        // ── Public API ────────────────────────────────────────────────────────

        /// <summary>Used by InjectPendingCharts for clubbed mode.</summary>
        public static void InjectChartForSheet(
            string xlsxPath,
            string sheetName,
            List<ResponseTimeRecord> records)
        {
            // records already sorted by avg desc from AppendToPackage
            InjectAllCharts(xlsxPath, records);
        }

        /// <summary>
        /// Adds a "Latency Charts" sheet with one mini chart per transaction
        /// sorted by Average descending, then saves to <paramref name="xlsxPath"/>.
        /// </summary>
        public static void AddMiniChartsAndSave(
            ExcelPackage package,
            List<ResponseTimeRecord> records,
            string sheetName,
            string xlsxPath)
        {
            // Charts sorted by Average descending (slowest first)
            var byAvg = new List<ResponseTimeRecord>(records);
            byAvg.Sort((a, b) => b.Average.CompareTo(a.Average));

            int n = byAvg.Count;

            // Compute axis scale — cap at 60s, values beyond shown at 65 with real label
            // Use same approach as JTL: fixed 0-70 scale, 70 label hidden
            var cs = package.Workbook.Worksheets.Add(sheetName);
            cs.Column(1).Width = 42;

            // Row 1 — title
            cs.Row(1).Height = TitleRowHt;
            cs.Cells[1, 1].Value = "Transaction Latency \u2013 Average vs 90th Percentile (Seconds)  |  Scale: 0 \u2013 60 s  (values >60 s shown capped at 65 s with actual label)";
            cs.Cells[1, 1].Style.Font.Bold = true;
            cs.Cells[1, 1].Style.Font.Size = 12;
            cs.Cells[1, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // Row 2 — scale row with coloured legend key in col A
            cs.Row(2).Height = ScaleRowHt;
            cs.Cells[2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            var rt = cs.Cells[2, 1].RichText;
            var rtScale = rt.Add("Scale    "); rtScale.Bold = true; rtScale.Color = System.Drawing.Color.Black;
            var rtAvgSq = rt.Add("\u25A0 "); rtAvgSq.Bold = true; rtAvgSq.Color = System.Drawing.Color.FromArgb(0x20, 0x6B, 0xA3);
            var rtAvgLb = rt.Add("Avg"); rtAvgLb.Bold = true; rtAvgLb.Color = System.Drawing.Color.Black;
            var rtSep = rt.Add("    "); rtSep.Color = System.Drawing.Color.Black;
            var rtP90Sq = rt.Add("\u25A0 "); rtP90Sq.Bold = true; rtP90Sq.Color = System.Drawing.Color.FromArgb(0xE3, 0x6C, 0x09);
            var rtP90Lb = rt.Add("P90"); rtP90Lb.Bold = true; rtP90Lb.Color = System.Drawing.Color.Black;

            // Register scale chart shell (chart1)
            var scaleChart = (OfficeOpenXml.Drawing.Chart.ExcelBarChart)
                cs.Drawings.AddChart("ScaleAxis",
                    OfficeOpenXml.Drawing.Chart.eChartType.BarClustered);
            scaleChart.SetPosition(1, 0, 1, 0);
            scaleChart.SetSize(1, 1);

            // Rows 3+ — one row per transaction, sorted by avg desc
            for (int i = 0; i < n; i++)
            {
                int row = 3 + i;
                cs.Row(row).Height = MiniRowHt;
                cs.Cells[row, 1].Value = byAvg[i].TransactionName;
                cs.Cells[row, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                var c = (OfficeOpenXml.Drawing.Chart.ExcelBarChart)
                    cs.Drawings.AddChart($"C{i}",
                        OfficeOpenXml.Drawing.Chart.eChartType.BarClustered);
                c.SetPosition(row - 1, 0, 1, 0);
                c.SetSize(1, 1);
            }

            cs.View.FreezePanes(3, 1);

            package.SaveAs(new FileInfo(xlsxPath));

            InjectAllCharts(xlsxPath, byAvg);
        }

        // ── ZIP injection ─────────────────────────────────────────────────────

        private static void InjectAllCharts(string xlsxPath, List<ResponseTimeRecord> records)
        {
            int n = records.Count;

            using var pkg = Package.Open(xlsxPath, FileMode.Open, FileAccess.ReadWrite);

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

            // Read rIds from existing drawing
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

            // chart[0] = scale axis
            if (chartParts.Count > 0)
            {
                var bytes = Encoding.UTF8.GetBytes(BuildScaleChartXml());
                using var s = chartParts[0].GetStream(FileMode.Create, FileAccess.Write);
                s.Write(bytes, 0, bytes.Length);
            }

            // chart[1..n] = transaction mini charts
            for (int i = 0; i < n && (i + 1) < chartParts.Count; i++)
            {
                var bytes = Encoding.UTF8.GetBytes(BuildMiniChartXml(records[i], i + 1));
                using var s = chartParts[i + 1].GetStream(FileMode.Create, FileAccess.Write);
                s.Write(bytes, 0, bytes.Length);
            }

            if (drawingPart != null)
            {
                var bytes = Encoding.UTF8.GetBytes(BuildDrawingXml(n, rIds));
                using var s = drawingPart.GetStream(FileMode.Create, FileAccess.Write);
                s.Write(bytes, 0, bytes.Length);
            }
        }

        // ── Scale chart XML ───────────────────────────────────────────────────

        private static string BuildScaleChartXml()
        {
            var sb = new StringBuilder(1200);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"");
            sb.Append(" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"");
            sb.Append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sb.Append("<c:lang val=\"en-US\"/><c:roundedCorners val=\"0\"/>");
            sb.Append("<c:chart><c:autoTitleDeleted val=\"1\"/><c:plotArea><c:layout/>");
            sb.Append("<c:barChart><c:barDir val=\"bar\"/><c:grouping val=\"clustered\"/>");
            sb.Append("<c:varyColors val=\"0\"/>");
            // Two invisible dummy series — force axis render + provide Avg/P90 legend colours
            foreach (var (idx, name) in new[] { (0, "Avg"), (1, "P90") })
            {
                sb.Append($"<c:ser><c:idx val=\"{idx}\"/><c:order val=\"{idx}\"/>");
                sb.Append($"<c:tx><c:v>{name}</c:v></c:tx>");
                sb.Append("<c:invertIfNegative val=\"0\"/>");
                sb.Append("<c:cat><c:strLit><c:ptCount val=\"1\"/>");
                sb.Append("<c:pt idx=\"0\"><c:v> </c:v></c:pt></c:strLit></c:cat>");
                sb.Append("<c:val><c:numLit><c:ptCount val=\"1\"/>");
                sb.Append("<c:pt idx=\"0\"><c:v>0</c:v></c:pt></c:numLit></c:val></c:ser>");
            }
            sb.Append("<c:gapWidth val=\"50\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>");
            sb.Append("<c:catAx><c:axId val=\"1\"/>");
            sb.Append("<c:scaling><c:orientation val=\"minMax\"/></c:scaling>");
            sb.Append("<c:delete val=\"1\"/><c:axPos val=\"l\"/><c:crossAx val=\"2\"/></c:catAx>");
            sb.Append("<c:valAx><c:axId val=\"2\"/>");
            sb.Append("<c:scaling><c:orientation val=\"minMax\"/>");
            sb.Append("<c:min val=\"0\"/><c:max val=\"70\"/></c:scaling>");
            sb.Append("<c:delete val=\"0\"/><c:axPos val=\"b\"/><c:majorGridlines/>");
            sb.Append("<c:numFmt formatCode=\"[&lt;70]0;[=70]&quot;&quot;;0\" sourceLinked=\"0\"/>");
            sb.Append("<c:tickLblPos val=\"low\"/>");
            sb.Append("<c:crossAx val=\"1\"/><c:crosses val=\"min\"/>");
            sb.Append("<c:crossBetween val=\"between\"/><c:majorUnit val=\"10\"/></c:valAx>");
            sb.Append("</c:plotArea>");
            sb.Append("<c:legend><c:delete val=\"1\"/></c:legend>");
            sb.Append("<c:plotVisOnly val=\"1\"/><c:dispBlanksAs val=\"zero\"/></c:chart>");
            sb.Append("<c:printSettings><c:headerFooter/>");
            sb.Append("<c:pageMargins b=\"0.25\" l=\"0.25\" r=\"0.25\" t=\"0.25\" header=\"0.3\" footer=\"0.3\"/>");
            sb.Append("<c:pageSetup/></c:printSettings></c:chartSpace>");
            return sb.ToString();
        }

        // ── Mini chart XML ────────────────────────────────────────────────────

        private static string BuildMiniChartXml(ResponseTimeRecord r, int idx)
        {
            const double CapAt = 65.0;

            double avgReal = System.Math.Round(r.Average, 3);
            double p90Real = r.Percentiles.TryGetValue("90% Line", out double p90v)
                ? System.Math.Round(p90v, 3) : 0;
            double avgBar = System.Math.Min(avgReal, CapAt);
            double p90Bar = System.Math.Min(p90Real, CapAt);

            string avgBarS = avgBar.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string p90BarS = p90Bar.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string avgLblS = avgReal.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string p90LblS = p90Real.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string a1 = (idx * 2 + 1).ToString();
            string a2 = (idx * 2 + 2).ToString();

            var sb = new StringBuilder(1500);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"");
            sb.Append(" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"");
            sb.Append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sb.Append("<c:lang val=\"en-US\"/><c:roundedCorners val=\"0\"/>");
            sb.Append("<c:chart><c:autoTitleDeleted val=\"1\"/><c:plotArea><c:layout/>");
            sb.Append("<c:barChart><c:barDir val=\"bar\"/><c:grouping val=\"clustered\"/>");
            sb.Append("<c:varyColors val=\"0\"/>");

            foreach (var (sidx, barVal, lblVal) in new[]
            {
                (0, avgBarS, avgLblS, "Avg"),
                (1, p90BarS, p90LblS, "P90"),
            }.Select(t => (t.Item1, t.Item2, t.Item3)))
            {
                sb.Append($"<c:ser><c:idx val=\"{sidx}\"/><c:order val=\"{sidx}\"/>");
                sb.Append($"<c:tx><c:v>{(sidx == 0 ? "Avg" : "P90")}</c:v></c:tx>");
                sb.Append("<c:invertIfNegative val=\"0\"/>");
                sb.Append("<c:dLbls><c:dLbl><c:idx val=\"0\"/>");
                sb.Append("<c:tx><c:rich><a:bodyPr/><a:lstStyle/>");
                sb.Append($"<a:p><a:r><a:t>{lblVal}</a:t></a:r></a:p></c:rich></c:tx>");
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
                sb.Append($"<c:pt idx=\"0\"><c:v>{barVal}</c:v></c:pt></c:numLit></c:val></c:ser>");
            }

            sb.Append("<c:gapWidth val=\"50\"/>");
            sb.Append($"<c:axId val=\"{a1}\"/><c:axId val=\"{a2}\"/></c:barChart>");
            sb.Append($"<c:catAx><c:axId val=\"{a1}\"/>");
            sb.Append("<c:scaling><c:orientation val=\"minMax\"/></c:scaling>");
            sb.Append("<c:delete val=\"1\"/><c:axPos val=\"l\"/>");
            sb.Append($"<c:crossAx val=\"{a2}\"/></c:catAx>");
            sb.Append($"<c:valAx><c:axId val=\"{a2}\"/>");
            sb.Append("<c:scaling><c:orientation val=\"minMax\"/>");
            sb.Append("<c:min val=\"0\"/><c:max val=\"70\"/></c:scaling>");
            sb.Append("<c:delete val=\"0\"/><c:axPos val=\"b\"/>");
            sb.Append("<c:tickLblPos val=\"none\"/>");
            sb.Append("<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>");
            sb.Append($"<c:crossAx val=\"{a1}\"/><c:crosses val=\"min\"/>");
            sb.Append("<c:crossBetween val=\"between\"/><c:majorUnit val=\"10\"/></c:valAx>");
            sb.Append("</c:plotArea>");
            sb.Append("<c:legend><c:delete val=\"1\"/></c:legend>");
            sb.Append("<c:plotVisOnly val=\"1\"/><c:dispBlanksAs val=\"zero\"/></c:chart>");
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

            int total = n + 1;
            for (int i = 0; i < total; i++)
            {
                string rId = i < rIds.Count ? rIds[i] : $"rId{i + 1}";
                long cy = i == 0 ? ScaleChartH : MiniChartH;
                int row = i + 1;
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
    }
}
