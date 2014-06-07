using System;
using System.Collections.Generic;
using System.Diagnostics;
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
            var timer = new Stopwatch();
            var timerAll = new Stopwatch();
            timerAll.Start();
            using (XLWorkbook wb = new XLWorkbook(XLEventTracking.Disabled))
            {
                using (var ws = wb.AddWorksheet("MergeCellsWorksheet"))
                {
                    int totalRows = 5000;

                    // Create some ranges
                    ws.Cell("AO1").Value = "A";
                    ws.Cell("AP1").Value = "B";
                    ws.Cell("AQ1").Value = "C";
                    ws.Cell("AR1").Value = "D";
                    ws.Cell("AS1").Value = "E";
                    ws.Cell("AT1").Value = "1";
                    ws.Cell("AU1").Value = "2";

                    var listRange = ws.Range("AO1:AU1");

                    timer.Start();
                    for (int i = 1; i <= totalRows; i++)
                    {
                        ws.Cell(i, 1).NewDataValidation.List(listRange);
                        Console.Clear();
                    }
                    timer.Stop();
                }

                wb.SaveAs(@"C:\temp\test.xlsx");
            }
            timerAll.Stop();
            Console.WriteLine();
            Console.WriteLine("Add validation Took {0}s", timer.Elapsed.TotalSeconds);
            Console.WriteLine("Complete Took {0}s", timerAll.Elapsed.TotalSeconds);
            Console.ReadKey();
        }

    }
}
