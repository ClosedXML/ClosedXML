using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace ClosedXML.Sandbox
{
    internal class PerformanceRunner
    {
        public static void TimeAction(Action action)
        {
            var stopwatch = Stopwatch.StartNew();
            action();
            Console.WriteLine("Action done in " + stopwatch.Elapsed);
        }

        private const int rowCount = 5000;

        public static void RunInsertTable()
        {
            var rows = new List<OneRow>();

            for (int i = 0; i < rowCount; i++)
            {
                var row = GenerateRow<OneRow>();
                rows.Add(row);
            }

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sheet 1");
            worksheet.Cell(1, 1).InsertTable(rows);

            CreateMergedCell(worksheet);

            worksheet.Columns().AdjustToContents();

            EmulateSave(workbook);
        }

        public static void OpenTestFile()
        {
            using (var wb = new XLWorkbook("test.xlsx"))
            {
                wb.RecalculateAllFormulas();
                var ws = wb.Worksheets.First();
                var cell = ws.FirstCellUsed();
                Console.WriteLine(cell.Value);
            }
        }

        private static void CreateMergedCell(IXLWorksheet worksheet)
        {
            worksheet.Cell(rowCount + 2, 1).Value = "Merged cell";
            var range = worksheet.Range(rowCount + 2, 1, rowCount + 2, 2);
            range.Row(1).Merge();
        }

        private static void EmulateSave(XLWorkbook workbook)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                workbook.SaveAs(memoryStream);
                memoryStream.Seek(0, SeekOrigin.Begin);
                Console.WriteLine("Total bytes = " + memoryStream.ToArray().Length);
            }
        }

        private static Random rnd = new Random();

        private static T GenerateRow<T>() where T : new()
        {
            var row = new T();

            var rowProps = row.GetType().GetProperties();

            var strings = rowProps.Where(p => p.PropertyType == typeof(string));
            var decimals = rowProps.Where(p => p.PropertyType == typeof(decimal));
            var ints = rowProps.Where(p => p.PropertyType == typeof(int) || p.PropertyType == typeof(int?));
            var dates = rowProps.Where(p => p.PropertyType == typeof(DateTime?));
            var timeSpans = rowProps.Where(p => p.PropertyType == typeof(TimeSpan?));
            var booleans = rowProps.Where(p => p.PropertyType == typeof(bool));

            // Format strings
            var tmpString = new StringBuilder();
            var tmpStringLength = rnd.Next(5, 50);
            for (int x = 0; x <= tmpStringLength; x++)
            {
                tmpString.Append((char)(rnd.Next(48, 120)));
            }
            foreach (var str in strings)
            {
                str.SetValue(row, tmpString.ToString(), null);
            }

            // Format decimals
            var tmpDec = (decimal)(rnd.Next(-10000, 100000) / (Math.Pow(10.0, rnd.Next(1, 4))));

            foreach (var dec in decimals)
            {
                dec.SetValue(row, tmpDec, null);
            }

            // Format ints
            var tmpInt = rnd.Next(-1000, 10000);

            foreach (var intValue in ints)
            {
                intValue.SetValue(row, tmpInt, null);
            }

            // Format dates
            var tmpDate = new DateTime(2012, 1, 1, 1, 1, 1);
            tmpDate = tmpDate.AddSeconds(rnd.Next(-10000, 100000));
            foreach (var dt in dates)
            {
                dt.SetValue(row, tmpDate, null);
            }

            // Format timespans
            var tmpTimespan = new TimeSpan(rnd.Next(1, 24), rnd.Next(1, 60), rnd.Next(1, 60));

            foreach (var ts in timeSpans)
            {
                ts.SetValue(row, tmpTimespan, null);
            }

            // Format booleans
            var tmpBool = (rnd.Next(0, 2) > 0);
            foreach (var bl in booleans)
            {
                bl.SetValue(row, tmpBool, null);
            }

            return row;
        }

        public static void PerformHeavyCalculation()
        {
            int rows = 200;
            int columns = 200;
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.Worksheets.Add("TestSheet");
                var lastColumnLetter = sheet.Column(columns).ColumnLetter();
                for (int i = 1; i <= rows; i++)
                {
                    for (int j = 1; j <= columns; j++)
                    {
                        if (i == 1)
                        {
                            sheet.Cell(i, j).FormulaA1 = string.Format("=ROUND({0}*SIN({0}),2)", j);
                        }
                        else
                        {
                            sheet.Cell(i, j).FormulaA1 = string.Format("=SUM({0}$1:{0}{1})/SUM($A{1}:${2}{1})",
                                sheet.Column(j).ColumnLetter(), i - 1, lastColumnLetter); // i.e. for K8 there will be =SUM(K$1:K7)/SUM($A7:$GR7)
                        }
                    }
                }

                var cells = sheet.CellsUsed();
                var sum1 = cells.Sum(cell => (double)cell.Value);
                Console.WriteLine("Total sum: {0:N2}", sum1);
            }
        }
    }
}
