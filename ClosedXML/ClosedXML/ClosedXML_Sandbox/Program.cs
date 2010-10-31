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


            //wb.Load(@"c:\Initial.xlsx");
            wb.SaveAs(@"C:\Excel Files\ForTesting\Sandbox.xlsx");
            //Console.ReadKey();
        }
        
        // Invalidate range references when they point to a deleted range.

        // Implement formulas
        // Implement grouping of rows and columns
        // Autosize rows/columns 
        // Save defaults to a .config file

        // Add/Copy/Paste (maybe another name?) rows, columns, ranges into an area.
    }
}
