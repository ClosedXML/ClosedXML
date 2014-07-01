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
            var wbSource = new XLWorkbook(@"c:\temp\original.xlsm");
            var wbTarget = new XLWorkbook();

            foreach (var ws in wbSource.Worksheets)
            {
                wbTarget.AddWorksheet(ws);
            }

            foreach (var r in wbSource.NamedRanges)
            {
                wbTarget.NamedRanges.Add(r.Name, r.Ranges);
            }

            wbTarget.SaveAs(@"c:\temp\saved.xlsm");
            Console.WriteLine("Done");
            //Console.ReadLine();
        }
    }
}
