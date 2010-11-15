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

            var lst = new List<Person>();
            lst.Add(new Person(){ Name = "Manuel", Age = 33});
            lst.Add(new Person() { Name = "Carlos", Age = 32 });

            ws.Cell(1, 1).Value = lst;
                        
            //wb.Load(@"c:\Initial.xlsx");
            wb.SaveAs(@"C:\Excel Files\ForTesting\Sandbox.xlsx");
            //Console.ReadKey();
        }

        class Person
        {
            public String Name { get; set; }
            public Int32 Age { get; set; }
        }

        // Save defaults to a .config file

        // Add/Copy/Paste (maybe another name?) rows, columns, ranges into an area.
    }
}
