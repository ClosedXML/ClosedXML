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
            var ws = wb.Worksheets.Add("Test");
            ws.Row(1).Style.Fill.BackgroundColor = Color.Red;
            ws.Cell(1, 1).Value = "Hello";

            // Also test painting a row/column, setting the value of a cell, and then moving it.
            // Change Internal references on XLRows/XLColumns so they return the values from Worksheet.Internal.Rows/Columns collection

            //wb.Load(@"c:\Initial.xlsx");
            wb.SaveAs(@"C:\Excel Files\ForTesting\Sandbox.xlsx");
            //Console.ReadKey();
        }
        
        // Apply a style to the entire sheet (not just the used cells)
        // Implement formulas
        // Implement grouping of rows and columns
        // Adjust rows/columns heights/widths
    }
}
