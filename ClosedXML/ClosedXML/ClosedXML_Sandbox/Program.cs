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

            var wb = new XLWorkbook(@"c:\temp\ColorBook.xlsx");
            var ws = wb.Worksheet("Sheet1");
            var fontColor = ws.Cell("C9").Style.Font.FontColor;
            Console.WriteLine("The result for cell is " + fontColor.Color.ToString());
            wb.SaveAs(@"c:\temp\saved.xlsx");
            Console.WriteLine("Done");
            Console.ReadLine();
        }
    }
}
