using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System.Xml;

namespace TestApp
{
    public static class ResponseTimeConverterExcelCharts
    {
        // ─────────────────────────────────────────────────────────────────────
        // Public API (called by ResponseTimeConverter)
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>
        /// Adds a new worksheet containing a horizontal clustered-bar chart
        /// that plots each percentile series for every transaction.
        /// </summary>
        public static void CreateChartSheet(
            ExcelPackage package,
            ExcelWorksheet dataSheet,
            List<ResponseTimeRecord> records,
            List<string> percentileHeaders,
            int percentileStartColumn,
            string sheetName = "Latency Charts")
        {
            var chartSheet = package.Workbook.Worksheets.Add(sheetName);

            var chart = (ExcelBarChart)chartSheet.Drawings.AddChart("LatencyChart", eChartType.BarClustered);
            chart.Title.Text = "Latency Percentile Comparison";
            chart.GapWidth = 120;
            chart.XAxis.MajorUnit = 1;

            int recordCount = records.Count;
            int lastRow = recordCount + 1;

            // ── Add one series per percentile column ──────────────────────────
            // ── Add Average series (always column 3) ─────────────────────────
            var avgSeries = (ExcelBarChartSerie)chart.Series.Add(
            dataSheet.Cells[2, 3, lastRow, 3],
            dataSheet.Cells[2, 1, lastRow, 1]);

            avgSeries.Header = "Average";
            avgSeries.DataLabel.ShowValue = true;
            avgSeries.DataLabel.Position = eLabelPosition.OutEnd;

            // ── Add 90th percentile series only ──────────────────────────────
            int p90Index = percentileHeaders.IndexOf("90% Line");
            if (p90Index >= 0)
            {
                int p90Col = percentileStartColumn + p90Index;
                var p90Series = (ExcelBarChartSerie)chart.Series.Add(
                dataSheet.Cells[2, p90Col, lastRow, p90Col],
                dataSheet.Cells[2, 1, lastRow, 1]);

                p90Series.Header = "90th Percentile";
                p90Series.DataLabel.ShowValue = true;
                p90Series.DataLabel.Position = eLabelPosition.OutEnd;
               
            }

            // ── Outlier-resistant axis maximum ────────────────────────────────
            double? axisMax = ComputeAxisMax(records);

            // ── Size and position ─────────────────────────────────────────────
            int chartHeight = Math.Max(500, recordCount * 40 + 100);
            chart.SetPosition(1, 0, 1, 0);
            chart.SetSize(1600, chartHeight);

            // ── Fix axis orientation via raw XML ──────────────────────────────
            FixBarChartAxisOrientation(chart, axisMax);
        }

        // ─────────────────────────────────────────────────────────────────────
        // Axis max computation
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>
        /// Returns an outlier-resistant axis maximum, or <c>null</c> when the
        /// data contains no positive values or no capping is needed.
        /// <para>
        /// Collects all percentile values, sorts them, takes the 75th percentile
        /// of those values, then uses 1.5× that as the chart max — but only when
        /// the true maximum is more than 3× the p75 value.
        /// </para>
        /// </summary>
        private static double? ComputeAxisMax(List<ResponseTimeRecord> records)
        {
            // Only consider Average and 90th percentile values for axis scaling
            var allValues = records
                .SelectMany(r =>
                {
                    var vals = new List<double> { r.Average };
                    if (r.Percentiles.TryGetValue("90% Line", out double p90))
                        vals.Add(p90);
                    return vals;
                })
                .Where(v => v > 0.0001)
                .OrderBy(v => v)
                .ToList();

            if (allValues.Count == 0)
                return null;

            double p75 = allValues[(int)(allValues.Count * 0.75)];
            double hardMax = allValues[^1];

            // Only cap the axis if the outlier is more than 3× the p75 value
            if (hardMax > p75 * 3)
                return Math.Ceiling(p75 * 3 * 10) / 10; // round up to 1 decimal

            return null;
        }

        // ─────────────────────────────────────────────────────────────────────
        // Chart XML patching
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>
        /// Directly patches the chart XML to produce a horizontal bar chart where:
        /// <list type="bullet">
        ///   <item>Category axis (transaction names) runs top-to-bottom (row 1 at top).</item>
        ///   <item>Value axis (numbers) runs left-to-right with 0 on the left.</item>
        /// </list>
        /// EPPlus's built-in <c>Orientation</c>/<c>Crosses</c> properties do not write
        /// the correct OOXML combination reliably, so we write the nodes ourselves.
        /// </summary>
        private static void FixBarChartAxisOrientation(ExcelChart chart, double? axisMax = null)
        {
            var xml = chart.ChartXml;
            var ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            // ── Category axis (catAx) ─────────────────────────────────────────
            // orientation = maxMin  → reverses order so row 1 appears at the top
            // crosses     = max     → value axis line sits at the bottom of the
            //                         reversed axis
            // tickLblPos  = low     → labels appear on the LEFT side
            var catAx = xml.SelectSingleNode("//c:catAx", ns);
            if (catAx != null)
            {
                SetOrCreateChildVal(xml, ns, catAx, "c:scaling/c:orientation", "maxMin");
                SetOrCreateChildVal(xml, ns, catAx, "c:crosses", "min");
                SetOrCreateChildVal(xml, ns, catAx, "c:tickLblPos", "low");
            }

            // ── Value axis (valAx) ────────────────────────────────────────────
            // orientation = minMax  → 0 on left, max on right (bars grow rightward)
            // crossesAt   = 0       → category axis intersects at value 0 (left edge)
            // tickLblPos  = low     → number labels appear at the bottom
            var valAx = xml.SelectSingleNode("//c:valAx", ns);
            if (valAx != null)
            {
                SetOrCreateChildVal(xml, ns, valAx, "c:scaling/c:min", "0");
                SetOrCreateChildVal(xml, ns, valAx, "c:scaling/c:orientation", "minMax");

                if (axisMax.HasValue)
                    SetOrCreateChildVal(
                        xml, ns, valAx,
                        "c:scaling/c:max",
                        axisMax.Value.ToString("G", System.Globalization.CultureInfo.InvariantCulture));

                SetOrCreateChildVal(xml, ns, valAx, "c:tickLblPos", "low");
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        // XML node helper
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>
        /// Walks (and creates, if absent) each element in <paramref name="relPath"/>
        /// relative to <paramref name="parent"/>, then sets its <c>val</c> attribute
        /// to <paramref name="val"/>.
        /// </summary>
        /// <param name="xml">The owning <see cref="XmlDocument"/>.</param>
        /// <param name="ns">Namespace manager with the "c" prefix registered.</param>
        /// <param name="parent">The node to start traversal from.</param>
        /// <param name="relPath">
        ///     Forward-slash-separated path of prefixed element names,
        ///     e.g. <c>"c:scaling/c:orientation"</c>.
        /// </param>
        /// <param name="val">Value to assign to the leaf element's <c>val</c> attribute.</param>
        private static void SetOrCreateChildVal(
            XmlDocument xml,
            XmlNamespaceManager ns,
            XmlNode parent,
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
                    // Strip the namespace prefix to obtain the local element name
                    var localName = part.Contains(':') ? part.Split(':')[1] : part;
                    child = xml.CreateElement("c", localName, chartNs);
                    node.AppendChild(child);
                }
                node = child;
            }

            // node is now the leaf element — set or overwrite the val attribute
            if (node.Attributes == null) return;

            var attr = node.Attributes["val"] ?? xml.CreateAttribute("val");
            attr.Value = val;
            node.Attributes.Append(attr);
        }
    }
}