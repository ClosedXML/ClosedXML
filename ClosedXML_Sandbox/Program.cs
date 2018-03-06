using ClosedXML.Excel;
using System;
using System.IO;

namespace ClosedXML_Sandbox
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            using (var stream = new MemoryStream())
            {
                // Create some test data to parse
                using (var book = new XLWorkbook())
                {
                    var sheet = book.AddWorksheet("Sheet 1");
                    sheet.Cell(1, 1).Value = "test1";
                    // Cell(1,2) is empty
                    sheet.Cell(1, 3).Value = null;
                    book.SaveAs(stream);
                }

                // Now load it back
                stream.Position = 0;
                using (var book = new XLWorkbook(stream))
                {
                    var sheet = book.Worksheets.Worksheet(1);
                    var result = sheet.Cell(1, 1).Value;
                    //Assert.AreEqual("test1", result);
                    result = sheet.Cell(1, 2).Value;
                    /* Assert.AreEqual(null, result); */ // Should be null!
                    result = sheet.Cell(1, 3).Value;
                    /* Assert.AreEqual(null, result);*/  // Should be null!
                }
                //Console.WriteLine("Running {0}", nameof(PerformanceRunner.OpenTestFile));
                //PerformanceRunner.TimeAction(PerformanceRunner.OpenTestFile);
                //Console.WriteLine();

                //Console.WriteLine("Running {0}", nameof(PerformanceRunner.RunInsertTable));
                //PerformanceRunner.TimeAction(PerformanceRunner.RunInsertTable);
                //Console.WriteLine();

                //Console.WriteLine("Press any key to continue");
                //Console.ReadKey();
            }
        }
    }
}
