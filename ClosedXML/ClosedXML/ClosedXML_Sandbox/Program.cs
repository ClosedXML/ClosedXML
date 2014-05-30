using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ClosedXML;
using System.Drawing;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
namespace ClosedXML_Sandbox
{
    class Program
    {
        private static void Main(string[] args)
        {
            var dt = new System.Data.DataTable();

            dt.Columns.Add("col1", typeof(string));
            dt.Columns.Add("col2", typeof(string));
            dt.Columns.Add("col3", typeof(double));

            var col1 = new string[] { "col1_val1", "col1_val2", "col1_val3" };
            var col2 = new string[] { "col2_val1", "col2_val2", "col2_val3" };

            var rnd = new Random();
            for (int i = 0; i < 10; i++)
            {
                var row = dt.NewRow();
                row["col1"] = col1[rnd.Next(0, 3)];
                row["col2"] = col2[rnd.Next(0, 3)];
                row["col3"] = rnd.NextDouble() * rnd.Next(10, 100);
                dt.Rows.Add(row);
            }

            var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("Sheet1");

            var table = sheet.Cell(1, 1).InsertTable(dt, "Table1", true);

            var range = table.DataRange;
            var header = sheet.Range(1, 1, 1, dt.Columns.Count);
            var dataRange = sheet.Range(header.FirstCell(), range.LastCell());

            var ptSheet = workbook.Worksheets.Add("Sheet2");

            var pt = ptSheet.PivotTables.AddNew("TablePivot", ptSheet.Cell(1, 1), dataRange);

            // COL2 then COL1
            pt.RowLabels.Add("col2");
            pt.RowLabels.Add("col1");

            pt.Values.Add("col3");

            workbook.SaveAs(@"c:\temp\saved.xlsx");
        }

        static void MainX(string[] args)
        {
            DateTime start, end;
            var times = new List<Double>();
            //foreach (var i in Enumerable.Range(1,10) )
            //{
                using (var wb = new XLWorkbook(@"c:\temp\test.xlsx"))
                {
                    start = DateTime.Now;
                    wb.SaveAs(@"c:\temp\saved.xlsx");
                    end = DateTime.Now;
                    var total = (end - start).TotalSeconds;
                    Console.WriteLine(total);
                    times.Add(total);
                }
            //}
            Console.WriteLine("Average: " + times.Average());
            Console.WriteLine("Done");
            Console.ReadKey();
        }
    }

    class PivotTableScenarios
    {
        public void RunAll(XLWorkbook wb)
        {
            Add_Row_Labels_and_Sum(wb);
            Add_category_on_row_and_SubCategory_on_column(wb); //not working
        }

        private void Add_Row_Labels_and_Sum(XLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("Add_Row_Labels_and_Sum");

            ws.Cell("A1").Value = "Category";
            ws.Cell("A2").Value = "A";
            ws.Cell("A3").Value = "B";
            ws.Cell("A4").Value = "B";

            ws.Cell("B1").Value = "SubCategory";
            ws.Cell("B2").Value = "X";
            ws.Cell("B3").Value = "Y";
            ws.Cell("B4").Value = "Z";

            ws.Cell("C1").Value = "Number";
            ws.Cell("C2").Value = 100;
            ws.Cell("C3").Value = 150;
            ws.Cell("C4").Value = 75;

            var pivotTable = ws.Range("A1:C4").CreatePivotTable(ws.Cell("E1"));
            pivotTable.RowLabels.Add("Category");
            pivotTable.RowLabels.Add("SubCategory");
            pivotTable.Values.Add("Number").SetSummaryFormula(XLPivotSummary.Sum);


        }

        private void Add_category_on_row_and_SubCategory_on_column(XLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("cat_on_row_SubCat_on_col");

            ws.Cell("A1").Value = "Category";
            ws.Cell("A2").Value = "A";
            ws.Cell("A3").Value = "B";
            ws.Cell("A4").Value = "B"
;
            ws.Cell("B1").Value = "SubCategory";
            ws.Cell("B2").Value = "X";
            ws.Cell("B3").Value = "Y";
            ws.Cell("B4").Value = "Z";

            ws.Cell("C1").Value = "Number";
            ws.Cell("C2").Value = 100;
            ws.Cell("C3").Value = 150;
            ws.Cell("C4").Value = 75;

            var pivotTable = ws.Range("A1:C4").CreatePivotTable(ws.Cell("E1"));
            pivotTable.RowLabels.Add("Category");
            pivotTable.ColumnLabels.Add("SubCategory");
            pivotTable.Values.Add("Number").SetSummaryFormula(XLPivotSummary.Sum);
        }
    }
}
