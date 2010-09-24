using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;

namespace ClosedXML_Sandbox
{
    class Program
    {
        static void Main(string[] args)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Print Areas");

            // Column Collection examples
            // Row Collection examples
            // Sheets examples
            // SheetTab examples
            
            // Add List<IXLRange> Ranges(...) to IXLRandge

            //Apply a style to the entire sheet (not just the used cells)

            wb.SaveAs(@"c:\Sandbox.xlsx");
            //Console.ReadKey();

        }
    }
}
