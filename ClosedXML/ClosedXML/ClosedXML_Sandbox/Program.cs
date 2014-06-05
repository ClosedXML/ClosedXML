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
            var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Test");
            worksheet.Cell(2, 2).SetValue("Text");
            var cf = worksheet.Cell(2, 2).AddConditionalFormat();
            var style = cf.WhenNotBlank();
            style
                    .Fill.SetBackgroundColor(XLColor.Red)
                    .Border.SetOutsideBorder(XLBorderStyleValues.Thick)
                    .Border.SetOutsideBorderColor(XLColor.Blue);
            workbook.SaveAs(@"C:\temp\saved.xlsx");
            Console.WriteLine("Done");
            //Console.ReadKey();
        }

    }
}
