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
            
            using (XLWorkbook wb = new XLWorkbook(XLEventTracking.Disabled))
            {
                using (var ws = wb.AddWorksheet("MergeCellsWorksheet"))
                {
                    int total = 5000;
                    var timer = System.Diagnostics.Stopwatch.StartNew();
                    // Insert some values
                    for (int i = 1; i <= total; i += 2)
                    {
                        ws.Cell(i, 1).Value = "Merge Cell 1";
                        ws.Cell(i, 3).Value = "Merge Cell 2";
                        ws.Cell(i, 4).Value = "Merge Cell 3";
                        ws.Cell((i + 1), 4).Value = "Merge Cell 4";

                        ws.Range("A" + i + ":B" + (i + 1)).Merge(false);
                        ws.Range("C" + i + ":C" + (i + 1)).Merge(false);
                        ws.Range("D" + i + ":E" + i).Merge(false);
                        ws.Range("D" + (i + 1) + ":E" + (i + 1)).Merge(false);
                    }
                    timer.Stop();
                    Console.WriteLine("Took {0}s", timer.Elapsed.TotalSeconds);
                }

                wb.SaveAs(@"c:\temp\saved.xlsx");
            }
            Console.WriteLine("Done");
            Console.ReadKey();
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
            //Console.ReadKey();
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
