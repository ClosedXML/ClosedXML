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
            var ws = wb.AddWorksheet("Sheet");
            ws.FirstCell().SetValue(1)
                .CellBelow().SetFormulaA1("IF(A1>0,Yes,No)") // Invalid
                .CellBelow().SetFormulaA1("IF(A1>0,\"Yes\",\"No\")") // OK
                .CellBelow().SetFormulaA1("IF(A1>0,TRUE,FALSE)"); // OK
            wb.SaveAs(@"c:\temp\saved.xlsx");
            Console.WriteLine("Done");
        }
    }
}
