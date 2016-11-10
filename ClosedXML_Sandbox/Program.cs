using ClosedXML.Excel;
using System.Linq;
using System;

namespace ClosedXML_Sandbox
{
    class Program
    {
        private static void Main(string[] args)
        {
            var path = "tmp.xlsx";
            using (var workbook = new XLWorkbook(path))
            {
                var ws1 = workbook.Worksheets.Last();

                int rowCount = 20; //example
                                   //first row headers
                for (int i = 1; i < rowCount; i++)
                {
                    var row = ws1.Row(i);
                    var values = row.Cells().Select(c => c.Value).ToArray();
                    IXLCell cell = row.Cell("E");     //.Cell(5);
                    bool isEmpty = cell.IsEmpty();      //always empty on this column, but in original have 6 lines
                    var val = cell.Value;   //cell.
                                            //vals += val + Environment.NewLine;
                }
            }

            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
        }
    }
}
