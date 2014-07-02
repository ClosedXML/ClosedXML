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
            
            var wbTarget = new XLWorkbook();
            var ws = wbTarget.AddWorksheet("Sheet1");
            ws.FirstCell().Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1);
            wbTarget.SaveAs(@"c:\temp\saved.xlsx");
            Console.WriteLine("Done");
            //Console.ReadLine();
        }
    }
}
