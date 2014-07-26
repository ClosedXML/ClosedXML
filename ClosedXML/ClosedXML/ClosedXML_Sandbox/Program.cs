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
            var wb = new XLWorkbook(@"c:\temp\sopos.xlsx");
            var ws = wb.Worksheet("3");
            Console.WriteLine(ws.Cell("E22").Value);
            //wb.SaveAs(@"c:\temp\saved.xlsx");
            Console.WriteLine("Done");
            Console.ReadKey();
        }
    }
}
