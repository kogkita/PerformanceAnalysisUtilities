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

            var (records, percentileHeaders) = ReadCsv(csvPath);

            using var package = new ExcelPackage();

            var dataSheet = WriteResponseSheet(package, records, percentileHeaders);

            if (records.Count > 0)
                CreateChartSheet(package, dataSheet, records.Count, percentileHeaders);

            package.SaveAs(new FileInfo(excelPath));
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

        private static ExcelWorksheet WriteResponseSheet(
            ExcelPackage package,
            List<ResponseTimeRecord> records,
            List<string> percentileHeaders)
        {
            var sheet = package.Workbook.Worksheets.Add("Response Times");

            int col = 1;

            sheet.Cells[1, col++].Value = "Transaction Name";
            sheet.Cells[1, col++].Value = "# Samples";
            sheet.Cells[1, col++].Value = "Average (Seconds)";
            sheet.Cells[1, col++].Value = "Median (Seconds)";

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
                sheet.Cells[row, col++].Value = r.ErrorPercent;

                row++;
            }

            sheet.Cells.AutoFitColumns();
            return sheet;
        }

        private static void CreateChartSheet(
            ExcelPackage package,
            ExcelWorksheet dataSheet,
            int recordCount,
            List<string> percentileHeaders)
        {
            var chartSheet = package.Workbook.Worksheets.Add("Latency Charts");
            var chart = chartSheet.Drawings.AddChart("LatencyChart", eChartType.ColumnClustered);

            chart.Title.Text = "Latency Percentile Comparison";

            int lastRow = recordCount + 1;
            int percentileStartColumn = 5;

            for (int i = 0; i < percentileHeaders.Count; i++)
            {
                int col = percentileStartColumn + i;

                var series = chart.Series.Add(
                    dataSheet.Cells[2, col, lastRow, col],
                    dataSheet.Cells[2, 1, lastRow, 1]);

                series.Header = percentileHeaders[i].Replace("% Line", "");
            }

            chart.SetPosition(1, 0, 1, 0);
            chart.SetSize(900, 500);
        }
    }
}