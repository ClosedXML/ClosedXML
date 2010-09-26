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
            wb.Load(@"c:\Initial.xlsx");
            wb.SaveAs(@"c:\Initial_Saved.xlsx");
            //Console.ReadKey();

        }
        // Modify IXLRange to have the "IXLRanges Ranges(...)" methods
        // Modify DefiningRanges example to show how to select multiple ranges
        // Apply a style to the entire sheet (not just the used cells)
        // Implement formulas
        // Implement grouping of rows and columns
    }
}
