using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;
using System.IO;

namespace ClosedXML_Sandbox
{
    class Program
    {
        static void Main(string[] args)
        {
            var wb = new XLWorkbook(@"C:\Excel Files\ForTesting\Sandbox.xlsx");
            var sheet1 = wb.Worksheets.Worksheet("Sheet1");
            sheet1.Cell(5, 1).Value = 200;
            sheet1.Cell(6, 1).Value = 200;
            sheet1.Cell(4, 1).FormulaA1 = "A2 + 5";

            var sheet2 = wb.Worksheets.Worksheet("Sheet2");
            sheet2.Cell("B3").Value = 50;

            wb.SaveAs(@"C:\Excel Files\ForTesting\Sandbox_Saved.xlsx");
            //File.Copy(@"C:\Excel Files\ForTesting\NamedRanges.xlsx", @"C:\Excel Files\ForTesting\Sandbox_Merged.xlsx", true);
            //var wb = new XLWorkbook(@"C:\Excel Files\ForTesting\NamedRanges.xlsx");
            //wb.Worksheets.Worksheet(0).Cell(1, 1).Value = "XXX";
            //var ws = wb.Worksheets.Add("Testing");
            //ws.PageSetup.PrintAreas.Add("A1:C3");
            //ws.Range("A1").CreateNamedRange("SuperTest");
            //ws.Cell(1, 1).Value = "Nada";
            //ws.Cell(1, 1).Style.Fill.BackgroundColor = Color.Red;
            //wb.NamedRanges.Delete("PeopleData");
            //wb.NamedRanges.Add("PeopleData", ws.Range("A1"), "SuperComment");
            //ws.Cell(1, 2).FormulaA1 = "1+1";

            //wb.MergeInto(@"C:\Excel Files\ForTesting\Sandbox_Merged.xlsx");
            //wb.SaveAs(@"C:\Excel Files\ForTesting\Sandbox_Saved.xlsx");
            //wb.SaveChangesTo(@"C:\Excel Files\ForTesting\Sandbox_Changes.xlsx");
            
        }

        class Person
        {
            public String Name { get; set; }
            public Int32 Age { get; set; }
        }

        // Save defaults to a .config file
    }
}
