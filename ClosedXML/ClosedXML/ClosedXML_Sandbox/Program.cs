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
            var wb = new XLWorkbook(@"C:\Excel Files\ForTesting\Excel2003.xlsx");
            //var ws = wb.Worksheets.Worksheet("Sheet1");

            //ws.Cell(1, 1).Value = "something";
            
            wb.SaveAs(@"C:\Excel Files\ForTesting\Sandbox.xlsx");
            //Console.ReadKey();
        }

        class Person
        {
            public String Name { get; set; }
            public Int32 Age { get; set; }
        }

        // Save defaults to a .config file
    }
}
