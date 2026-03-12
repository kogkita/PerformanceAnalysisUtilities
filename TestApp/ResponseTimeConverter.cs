// ResponseTimeConverter.cs

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using OfficeOpenXml;

namespace TestApp
{
    public class ResponseTimeConverter
    {
        public void ConvertCsvToExcel(string csvFilePath, string excelFilePath, List<double> percentiles)
        {
            DataTable dataTable = new DataTable();
            // Read CSV file and populate DataTable
            using (var reader = new StreamReader(csvFilePath))
            {
                // Read headers
                var headers = reader.ReadLine().Split(',');
                foreach (var header in headers)
                {
                    dataTable.Columns.Add(header);
                }
                
                // Read rows
                while (!reader.EndOfStream)
                {
                    var rows = reader.ReadLine().Split(',');
                    DataRow row = dataTable.NewRow();
                    for (int i = 0; i < headers.Length; i++)
                    {
                        // Convert milliseconds to seconds and populate rows
                        if (double.TryParse(rows[i], out double milliseconds))
                        {
                            row[i] = milliseconds / 1000; // Convert to seconds
                        }
                        else
                        {
                            row[i] = rows[i];
                        }
                    }
                    dataTable.Rows.Add(row);
                }
            }
            
            // Create Excel file with dynamic percentile headers
            using (ExcelPackage excel = new ExcelPackage())
            {
                var worksheet = excel.Workbook.Worksheets.Add("ResponseTimeData");
                worksheet.Cells[1, 1].LoadFromDataTable(dataTable, true);

                // Add percentile columns
                for (int i = 0; i < percentiles.Count; i++)
                {
                    double percentileValue = CalculatePercentile(dataTable, percentiles[i]);
                    worksheet.Cells[1, dataTable.Columns.Count + i + 1].Value = percentiles[i] + "th Percentile";
                    worksheet.Cells[2, dataTable.Columns.Count + i + 1].Value = percentileValue;
                }

                // Save to file
                FileInfo excelFile = new FileInfo(excelFilePath);
                excel.SaveAs(excelFile);
            }
        }
        
        private double CalculatePercentile(DataTable dataTable, double percentile)
        {
            int count = dataTable.Rows.Count;
            if (count == 0) return 0;
            
            List<double> values = new List<double>();
            foreach (DataRow row in dataTable.Rows)
            {
                if (double.TryParse(row[0].ToString(), out double val))
                {
                    values.Add(val);
                }
            }
            values.Sort();
            int index = (int)Math.Ceiling(percentile / 100 * count) - 1;
            return values[Math.Max(0, Math.Min(index, values.Count - 1))];
        }
    }
}