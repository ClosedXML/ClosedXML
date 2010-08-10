using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.Excel.Style;

namespace ClosedXML_Sandbox
{
    class Program
    {
        static void Main(string[] args)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("New Sheet");
            ws.Cell(2, 2).Value = "Hello!";
            ws.Cell(2, 2).Style.Font.Bold = true;
            wb.SaveAs(@"c:\Sandbox.xlsx");
            //Console.ReadKey();
        }
    }
}
