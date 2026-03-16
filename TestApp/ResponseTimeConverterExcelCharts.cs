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
        private const long ChartW = 1400L * 9525L;
        private const long ScaleChartH = 55L * 9525L;
        private const long MiniChartH = 55L * 9525L;
        private const double TitleRowHt = 20.0;
        private const double ScaleRowHt = 42.0;
        private const double MiniRowHt = 42.0;

        // ── Public API ────────────────────────────────────────────────────────

        /// <summary>Single-file path: builds chart sheet, saves, injects XML.</summary>
        public static void AddMiniChartsAndSave(
            ExcelPackage package,
            List<ResponseTimeRecord> records,
            string sheetName,
            string xlsxPath)
        {
            var byAvg = records.OrderByDescending(r => r.Average).ToList();
            BuildChartSheetShells(package, sheetName, byAvg);
            package.SaveAs(new FileInfo(xlsxPath));
            InjectChartsForSheet(xlsxPath, sheetName, byAvg);
        }

        /// <summary>
        /// Creates the chart worksheet with EPPlus shell charts registered.
        /// Scale shell is registered FIRST (becomes chart1). Used by both paths.
        /// </summary>
        public static ExcelWorksheet BuildChartSheetShells(
            ExcelPackage package,
            string sheetName,
            List<ResponseTimeRecord> records)
        {
            int n = records.Count;
            var cs = package.Workbook.Worksheets.Add(sheetName);
            cs.Column(1).Width = 42;

            cs.Row(1).Height = TitleRowHt;
            cs.Cells[1, 1].Value = "Transaction Latency \u2013 Average vs 90th Percentile (Seconds)  |  Scale: 0 \u2013 60 s  (values >60 s capped at 65 s, actual value shown)";
            cs.Cells[1, 1].Style.Font.Bold = true;
            cs.Cells[1, 1].Style.Font.Size = 12;
            cs.Cells[1, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            cs.Row(2).Height = ScaleRowHt;
            cs.Cells[2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            var rt = cs.Cells[2, 1].RichText;
            var rScale = rt.Add("Scale    "); rScale.Bold = true; rScale.Color = System.Drawing.Color.Black;
            var rAvgSq = rt.Add("\u25A0 "); rAvgSq.Bold = true; rAvgSq.Color = System.Drawing.Color.FromArgb(0x20, 0x6B, 0xA3);
            var rAvgLb = rt.Add("Avg"); rAvgLb.Bold = true; rAvgLb.Color = System.Drawing.Color.Black;
            var rSep = rt.Add("    "); rSep.Color = System.Drawing.Color.Black;
            var rP90Sq = rt.Add("\u25A0 "); rP90Sq.Bold = true; rP90Sq.Color = System.Drawing.Color.FromArgb(0xE3, 0x6C, 0x09);
            var rP90Lb = rt.Add("P90"); rP90Lb.Bold = true; rP90Lb.Color = System.Drawing.Color.Black;

            // Scale shell FIRST → becomes chart1
            var scaleShell = (OfficeOpenXml.Drawing.Chart.ExcelBarChart)
                cs.Drawings.AddChart("ScaleAxis", OfficeOpenXml.Drawing.Chart.eChartType.BarClustered);
            scaleShell.SetPosition(1, 0, 1, 0);
            scaleShell.SetSize(1, 1);

            for (int i = 0; i < n; i++)
            {
                int row = 3 + i;
                cs.Row(row).Height = MiniRowHt;
                cs.Cells[row, 1].Value = records[i].TransactionName;
                cs.Cells[row, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                var c = (OfficeOpenXml.Drawing.Chart.ExcelBarChart)
                    cs.Drawings.AddChart($"C{i}", OfficeOpenXml.Drawing.Chart.eChartType.BarClustered);
                c.SetPosition(row - 1, 0, 1, 0);
                c.SetSize(1, 1);
            }
            cs.View.FreezePanes(3, 1);
            return cs;
        }

        /// <summary>Used by InjectPendingCharts for clubbed mode.</summary>
        public static void InjectChartForSheet(
            string xlsxPath,
            string sheetName,
            List<ResponseTimeRecord> records)
        {
            InjectChartsForSheet(xlsxPath, sheetName, records);
        }

        // ── ZIP injection (sheet-aware) ────────────────────────────────────────

        private static void InjectChartsForSheet(
            string xlsxPath,
            string sheetName,
            List<ResponseTimeRecord> records)
        {
            int n = records.Count;
            using var pkg = Package.Open(xlsxPath, FileMode.Open, FileAccess.ReadWrite);

            // Step 1: find sheet rId from workbook.xml by name
            var wbPart = pkg.GetPart(new Uri("/xl/workbook.xml", UriKind.Relative));
            string wbXml;
            using (var sr = new StreamReader(wbPart.GetStream(FileMode.Open, FileAccess.Read)))
                wbXml = sr.ReadToEnd();

            string? sheetRid = null;
            foreach (System.Text.RegularExpressions.Match m in
                System.Text.RegularExpressions.Regex.Matches(wbXml, @"<sheet\s[^>]*name=""([^""]+)""[^>]*/?>"))
            {
                if (m.Groups[1].Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    var ridM = System.Text.RegularExpressions.Regex.Match(m.Value, @"r:id=""([^""]+)""");
                    if (ridM.Success) sheetRid = ridM.Groups[1].Value;
                    break;
                }
            }
            if (sheetRid == null) return;

            // Step 2: resolve rId to worksheet part via workbook.xml.rels
            var wbRelsUri = new Uri("/xl/_rels/workbook.xml.rels", UriKind.Relative);
            if (!pkg.PartExists(wbRelsUri)) return;
            string wbRels;
            using (var sr = new StreamReader(pkg.GetPart(wbRelsUri).GetStream(FileMode.Open, FileAccess.Read)))
                wbRels = sr.ReadToEnd();

            var wsRelMatch = System.Text.RegularExpressions.Regex.Match(
                wbRels, $@"Id=""{System.Text.RegularExpressions.Regex.Escape(sheetRid)}""[^>]*Target=""([^""]+)""");
            if (!wsRelMatch.Success) return;

            string wsTarget = wsRelMatch.Groups[1].Value;
            string wsUriStr = wsTarget.StartsWith("/") ? wsTarget : "/xl/" + wsTarget;

            // Step 3: find drawing from worksheet rels
            var wsRelsUri = new Uri(
                wsUriStr.Replace("/xl/worksheets/", "/xl/worksheets/_rels/") + ".rels",
                UriKind.Relative);
            if (!pkg.PartExists(wsRelsUri)) return;
            string wsRels;
            using (var sr = new StreamReader(pkg.GetPart(wsRelsUri).GetStream(FileMode.Open, FileAccess.Read)))
                wsRels = sr.ReadToEnd();

            var drawingMatch = System.Text.RegularExpressions.Regex.Match(
                wsRels, @"Type=""[^""]*drawing[^""]*""\s+Target=""([^""]+)""");
            if (!drawingMatch.Success) return;

            string drawingTarget = drawingMatch.Groups[1].Value;
            string drawingUriStr;
            if (drawingTarget.StartsWith("/"))
                drawingUriStr = drawingTarget;
            else
            {
                var baseUri = new Uri("http://x/xl/worksheets/");
                drawingUriStr = "/" + new Uri(baseUri, drawingTarget).AbsolutePath.TrimStart('/');
            }
            if (!pkg.PartExists(new Uri(drawingUriStr, UriKind.Relative))) return;

            var drawingPart = pkg.GetPart(new Uri(drawingUriStr, UriKind.Relative));
            string drawingXml;
            using (var sr = new StreamReader(drawingPart.GetStream(FileMode.Open, FileAccess.Read)))
                drawingXml = sr.ReadToEnd();

            var chartPartIds = new List<string>();
            foreach (System.Text.RegularExpressions.Match m in
                System.Text.RegularExpressions.Regex.Matches(drawingXml, @"r:id=""([^""]+)"""))
                chartPartIds.Add(m.Groups[1].Value);

            if (chartPartIds.Count == 0) return;

            var drawingRelsUri = new Uri(
                drawingPart.Uri.ToString().Replace("/xl/drawings/", "/xl/drawings/_rels/") + ".rels",
                UriKind.Relative);
            if (!pkg.PartExists(drawingRelsUri)) return;

            string drawingRels;
            using (var sr = new StreamReader(pkg.GetPart(drawingRelsUri).GetStream(FileMode.Open, FileAccess.Read)))
                drawingRels = sr.ReadToEnd();

            var rIdToUri = new Dictionary<string, string>();
            foreach (System.Text.RegularExpressions.Match m in
                System.Text.RegularExpressions.Regex.Matches(
                    drawingRels, @"Id=""([^""]+)""\s+[^>]*Target=""([^""]+)"""))
                rIdToUri[m.Groups[1].Value] = m.Groups[2].Value;

            // Scale chart (index 0)
            if (chartPartIds.Count > 0 && rIdToUri.TryGetValue(chartPartIds[0], out var scaleUri))
            {
                string fullUri = scaleUri.StartsWith("/") ? scaleUri : "/xl/charts/" + System.IO.Path.GetFileName(scaleUri);
                if (pkg.PartExists(new Uri(fullUri, UriKind.Relative)))
                {
                    var bytes = Encoding.UTF8.GetBytes(BuildScaleChartXml());
                    using var s = pkg.GetPart(new Uri(fullUri, UriKind.Relative)).GetStream(FileMode.Create, FileAccess.Write);
                    s.Write(bytes, 0, bytes.Length);
                }
            }

            // Transaction charts
            for (int i = 0; i < n && (i + 1) < chartPartIds.Count; i++)
            {
                if (!rIdToUri.TryGetValue(chartPartIds[i + 1], out var txUri)) continue;
                string fullUri = txUri.StartsWith("/") ? txUri : "/xl/charts/" + System.IO.Path.GetFileName(txUri);
                if (!pkg.PartExists(new Uri(fullUri, UriKind.Relative))) continue;
                var bytes = Encoding.UTF8.GetBytes(BuildMiniChartXml(records[i], i + 1));
                using var s = pkg.GetPart(new Uri(fullUri, UriKind.Relative)).GetStream(FileMode.Create, FileAccess.Write);
                s.Write(bytes, 0, bytes.Length);
            }

            // Drawing XML
            var drawingBytes = Encoding.UTF8.GetBytes(BuildDrawingXml(n, chartPartIds));
            using var ds = drawingPart.GetStream(FileMode.Create, FileAccess.Write);
            ds.Write(drawingBytes, 0, drawingBytes.Length);
        }

        // ── Scale chart XML ───────────────────────────────────────────────────

        private static string BuildScaleChartXml()
        {
            var sb = new StringBuilder(1024);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"");
            sb.Append(" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"");
            sb.Append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sb.Append("<c:lang val=\"en-US\"/><c:roundedCorners val=\"0\"/>");
            sb.Append("<c:chart><c:autoTitleDeleted val=\"1\"/><c:plotArea><c:layout/>");
            sb.Append("<c:barChart><c:barDir val=\"bar\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>");
            foreach (var (idx, name) in new[] { (0, "Avg"), (1, "P90") })
            {
                sb.Append($"<c:ser><c:idx val=\"{idx}\"/><c:order val=\"{idx}\"/>");
                sb.Append($"<c:tx><c:v>{name}</c:v></c:tx>");
                sb.Append("<c:invertIfNegative val=\"0\"/>");
                sb.Append("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>");
                sb.Append("<c:cat><c:strLit><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v> </c:v></c:pt></c:strLit></c:cat>");
                sb.Append("<c:val><c:numLit><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>0</c:v></c:pt></c:numLit></c:val></c:ser>");
            }
            sb.Append("<c:gapWidth val=\"50\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>");
            sb.Append("<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling>");
            sb.Append("<c:delete val=\"1\"/><c:axPos val=\"l\"/><c:crossAx val=\"2\"/></c:catAx>");
            sb.Append("<c:valAx><c:axId val=\"2\"/>");
            sb.Append("<c:scaling><c:orientation val=\"minMax\"/><c:min val=\"0\"/><c:max val=\"70\"/></c:scaling>");
            sb.Append("<c:delete val=\"0\"/><c:axPos val=\"b\"/><c:majorGridlines/>");
            sb.Append("<c:numFmt formatCode=\"[&lt;70]0;[=70]&quot;&quot;;0\" sourceLinked=\"0\"/>");
            sb.Append("<c:tickLblPos val=\"low\"/><c:crossAx val=\"1\"/><c:crosses val=\"min\"/>");
            sb.Append("<c:crossBetween val=\"between\"/><c:majorUnit val=\"10\"/></c:valAx>");
            sb.Append("</c:plotArea><c:legend><c:delete val=\"1\"/></c:legend>");
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
            // Average already in seconds from ReadCsv
            double avgReal = System.Math.Round(r.Average, 3);
            double p90Real = r.Percentiles.TryGetValue("90% Line", out double p90v)
                ? System.Math.Round(p90v, 3) : 0;
            double avgBar = System.Math.Min(avgReal, CapAt);
            double p90Bar = System.Math.Min(p90Real, CapAt);
            string avgBarStr = avgBar.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string p90BarStr = p90Bar.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string avgLblStr = avgReal.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string p90LblStr = p90Real.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string safeName = EscapeXml(r.TransactionName);
            string a1 = (idx * 2 + 1).ToString();
            string a2 = (idx * 2 + 2).ToString();

            var sb = new StringBuilder(1200);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"");
            sb.Append(" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"");
            sb.Append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sb.Append("<c:lang val=\"en-US\"/><c:roundedCorners val=\"0\"/>");
            sb.Append("<c:chart><c:autoTitleDeleted val=\"1\"/><c:plotArea><c:layout/>");
            sb.Append("<c:barChart><c:barDir val=\"bar\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>");

            foreach (var (sidx, barVal, lblVal) in new[]
            {
                (0, avgBarStr, avgLblStr),
                (1, p90BarStr, p90LblStr)
            })
            {
                string serName = sidx == 0 ? "Avg" : "P90";
                sb.Append($"<c:ser><c:idx val=\"{sidx}\"/><c:order val=\"{sidx}\"/>");
                sb.Append($"<c:tx><c:v>{serName}</c:v></c:tx>");
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
                sb.Append($"<c:pt idx=\"0\"><c:v>{safeName}</c:v></c:pt></c:strLit></c:cat>");
                sb.Append($"<c:val><c:numLit><c:ptCount val=\"1\"/>");
                sb.Append($"<c:pt idx=\"0\"><c:v>{barVal}</c:v></c:pt></c:numLit></c:val></c:ser>");
            }

            sb.Append("<c:gapWidth val=\"50\"/>");
            sb.Append($"<c:axId val=\"{a1}\"/><c:axId val=\"{a2}\"/></c:barChart>");
            sb.Append($"<c:catAx><c:axId val=\"{a1}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling>");
            sb.Append("<c:delete val=\"1\"/><c:axPos val=\"l\"/>");
            sb.Append($"<c:crossAx val=\"{a2}\"/></c:catAx>");
            sb.Append($"<c:valAx><c:axId val=\"{a2}\"/>");
            sb.Append("<c:scaling><c:orientation val=\"minMax\"/><c:min val=\"0\"/><c:max val=\"70\"/></c:scaling>");
            sb.Append("<c:delete val=\"0\"/><c:axPos val=\"b\"/>");
            sb.Append("<c:tickLblPos val=\"none\"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>");
            sb.Append($"<c:crossAx val=\"{a1}\"/><c:crosses val=\"min\"/>");
            sb.Append("<c:crossBetween val=\"between\"/><c:majorUnit val=\"10\"/></c:valAx>");
            sb.Append("</c:plotArea><c:legend><c:delete val=\"1\"/></c:legend>");
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
                sb.Append($"<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{ChartW}\" cy=\"{cy}\"/></xdr:xfrm>");
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
