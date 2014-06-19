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
            var wb = new XLWorkbook();
            foreach (var sheetNum in Enumerable.Range(1, 1))
            {
                CreateSheet(wb, sheetNum);
                Console.WriteLine("Sheet " + sheetNum);
            }
            wb.SaveAs(@"c:\temp\saved.xlsx");
            Console.WriteLine("Done");
        }

        private static void CreateSheet(XLWorkbook wb, Int32 sheetNum)
        {
            using (var ws = wb.AddWorksheet("Sheet " + sheetNum))
            {
                foreach (var ro in Enumerable.Range(1, 1000))
                {
                    foreach (var co in Enumerable.Range(1, 100))
                    {
                        ws.Cell(ro, co).Value = ro + co;
                    }
                }
                ws.Dispose();
            }
        }

    }
}
