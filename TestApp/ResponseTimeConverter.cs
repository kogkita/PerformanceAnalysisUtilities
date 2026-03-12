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
        public double P90 { get; set; }
        public double P80 { get; set; }
        public double P70 { get; set; }
        public double Min { get; set; }
        public double Max { get; set; }
        public double ErrorPercent { get; set; }
    }

    public static class ResponseTimeConverter
    {
        public static void Convert(string csvPath, string excelPath)
        {
            var records = ReadCsv(csvPath);
            ExcelPackage.License.SetNonCommercialPersonal("Response Time Converter");
          //  ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage();

            var dataSheet = WriteResponseSheet(package, records);

            if (records.Count > 0)
                CreateChartSheet(package, dataSheet, records.Count);

            package.SaveAs(new FileInfo(excelPath));
        }

        private static List<ResponseTimeRecord> ReadCsv(string csvPath)
        {
            var records = new List<ResponseTimeRecord>();

            if (!File.Exists(csvPath))
                throw new FileNotFoundException("CSV file not found", csvPath);

            var lines = File.ReadAllLines(csvPath);

            for (int i = 1; i < lines.Length; i++)
            {
                var values = lines[i].Split(',');

                if (values.Length < 10)
                    continue;

                if (values[0].Trim().Equals("TOTAL", StringComparison.OrdinalIgnoreCase))
                    continue;

                records.Add(new ResponseTimeRecord
                {
                    TransactionName = values[0],
                    Samples = ParseInt(values[1]),
                    Average = ToSeconds(values[2]),
                    Median = ToSeconds(values[3]),
                    P90 = ToSeconds(values[4]),
                    P80 = ToSeconds(values[5]),
                    P70 = ToSeconds(values[6]),
                    Min = ToSeconds(values[7]),
                    Max = ToSeconds(values[8]),
                    ErrorPercent = ParsePercent(values[9])
                });
            }

            return records;
        }

        private static int ParseInt(string value)
        {
            if (int.TryParse(value, out int result))
                return result;

            return 0;
        }

        private static double ParsePercent(string value)
        {
            value = value.Replace("%", "");

            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
                return result;

            return 0;
        }

        private static double ToSeconds(string ms)
        {
            if (double.TryParse(ms, NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
                return value / 1000;

            return 0;
        }

        private static ExcelWorksheet WriteResponseSheet(ExcelPackage package, List<ResponseTimeRecord> records)
        {
            var sheet = package.Workbook.Worksheets.Add("Response Times");

            string[] headers =
            {
                "Transaction Name",
                "# Samples",
                "Average (Seconds)",
                "Median (Seconds)",
                "90th Percentile (Seconds)",
                "80th Percentile (Seconds)",
                "70th Percentile (Seconds)",
                "Min (Seconds)",
                "Max (Seconds)",
                "Error %"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                sheet.Cells[1, i + 1].Value = headers[i];
                sheet.Cells[1, i + 1].Style.Font.Bold = true;
                sheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            }

            int row = 2;

            foreach (var r in records)
            {
                sheet.Cells[row, 1].Value = r.TransactionName;
                sheet.Cells[row, 2].Value = r.Samples;
                sheet.Cells[row, 3].Value = r.Average;
                sheet.Cells[row, 4].Value = r.Median;
                sheet.Cells[row, 5].Value = r.P90;
                sheet.Cells[row, 6].Value = r.P80;
                sheet.Cells[row, 7].Value = r.P70;
                sheet.Cells[row, 8].Value = r.Min;
                sheet.Cells[row, 9].Value = r.Max;
                sheet.Cells[row, 10].Value = r.ErrorPercent;

                row++;
            }

            sheet.Cells.AutoFitColumns();

            return sheet;
        }

        private static void CreateChartSheet(ExcelPackage package, ExcelWorksheet dataSheet, int recordCount)
        {
            var chartSheet = package.Workbook.Worksheets.Add("Latency Charts");

            var chart = chartSheet.Drawings.AddChart("LatencyChart", eChartType.ColumnClustered);

            chart.Title.Text = "Latency Percentile Comparison";

            int lastRow = recordCount + 1;

            var p90 = chart.Series.Add(
                dataSheet.Cells[2, 5, lastRow, 5],
                dataSheet.Cells[2, 1, lastRow, 1]);
            p90.Header = "P90";

            var p80 = chart.Series.Add(
                dataSheet.Cells[2, 6, lastRow, 6],
                dataSheet.Cells[2, 1, lastRow, 1]);
            p80.Header = "P80";

            var p70 = chart.Series.Add(
                dataSheet.Cells[2, 7, lastRow, 7],
                dataSheet.Cells[2, 1, lastRow, 1]);
            p70.Header = "P70";

            chart.SetPosition(1, 0, 1, 0);
            chart.SetSize(900, 500);
        }
    }
}