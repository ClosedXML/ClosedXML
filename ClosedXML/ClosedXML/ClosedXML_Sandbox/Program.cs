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
            var ws = wb.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue("Hello")
                .CellBelow().SetValue("Hellos")
                .CellBelow().SetValue("Hell")
                .CellBelow().SetValue("Holl");

            ws.RangeUsed().AddConditionalFormat().WhenStartsWith("Hell")
                .Fill.SetBackgroundColor(XLColor.Red)
                .Border.SetOutsideBorder(XLBorderStyleValues.Thick)
                .Border.SetOutsideBorderColor(XLColor.Blue)
                .Font.SetBold();

            
            wb.SaveAs(@"c:\temp\saved.xlsx");
            Console.WriteLine("Done");
            Console.ReadLine();
        }
    }
}
