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
            var wb = new XLWorkbook();
            foreach (var wsNum in Enumerable.Range(0, 5))
            {
                wb.Worksheets.Add("Original Pos. is " + wsNum.ToString());
            }

            // Move first worksheet to the last position
            wb.Worksheets.Worksheet(0).SheetIndex = wb.Worksheets.Count();

            // Delete worksheet on position 2 (in this case it's where original position = 3)
            wb.Worksheets.Worksheet(2).Delete();

            // Swap sheets in positions 0 and 1
            wb.Worksheets.Worksheet(1).SheetIndex = 0;

            wb.SaveAs(@"C:\Excel Files\ForTesting\OrganizingSheets.xlsx");
        }
        static void xMain(string[] args)
        {
            List<Double> running = new List<Double>();
            foreach (Int32 r in Enumerable.Range(1, 1))
            {
                var startTotal = DateTime.Now;

                FillStyles();
                var wb = new XLWorkbook();
                foreach (var i in Enumerable.Range(1, 3))
                {
                    var ws = wb.Worksheets.Add("Sheet" + i);
                    foreach (var ro in Enumerable.Range(1, 100))
                    {
                        foreach (var co in Enumerable.Range(1, 100))
                        {
                            ws.Cell(ro, co).Style = GetRandomStyle();
                            ws.Cell(ro, co).Value = GetRandomValue();
                        }
                        //System.Threading.Thread.Sleep(10);
                    }
                }
                var start = DateTime.Now;
                wb.SaveAs(@"C:\Excel Files\ForTesting\Benchmark.xlsx");
                var end = DateTime.Now;
                Console.WriteLine("Saved in {0} secs.", (end - start).TotalSeconds);

                var start1 = DateTime.Now;
                var wb1 = new XLWorkbook(@"C:\Excel Files\ForTesting\Benchmark.xlsx");

                var end1 = DateTime.Now;
                Console.WriteLine("Loaded in {0} secs.", (end1 - start1).TotalSeconds);
                var start2 = DateTime.Now;
                wb1.SaveAs(@"C:\Excel Files\ForTesting\Benchmark_Saved.xlsx");
                var end2 = DateTime.Now;
                Console.WriteLine("Saved back in {0} secs.", (end2 - start2).TotalSeconds);

                var endTotal = DateTime.Now;
                Console.WriteLine("It all took {0} secs.", (endTotal - startTotal).TotalSeconds);
                running.Add((endTotal - startTotal).TotalSeconds);
            }
            Console.WriteLine("-------");
            Console.WriteLine("Avg total time: {0}", running.Average());
            //Console.ReadKey();
        }

        private static IXLStyle style1;
        private static IXLStyle style2;
        private static IXLStyle style3;
        private static void FillStyles()
        {

                    style1 = XLWorkbook.DefaultStyle;
                    style1.Font.Bold = true;
                    style1.Fill.BackgroundColor = Color.Azure;
                    style1.Border.BottomBorder = XLBorderStyleValues.Medium;
                    style1.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    style2 = XLWorkbook.DefaultStyle;
                    style2.Font.Italic = true;
                    style2.Fill.BackgroundColor = Color.Orange;
                    style2.Border.LeftBorder = XLBorderStyleValues.Medium;
                    style2.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    style3 = XLWorkbook.DefaultStyle;
                    style3.Font.FontColor = Color.Red;
                    style3.Fill.PatternColor = Color.Blue;
                    style3.Fill.PatternType = XLFillPatternValues.DarkTrellis;
                    style3.Border.DiagonalBorder = XLBorderStyleValues.Dotted;
        }
        private static IXLStyle GetRandomStyle()
        {
            
            var val = rnd.Next(1, 4);
            if (val == 1)
            {
                return style1;
            }
            else if (val == 2)
            {
                return style2;
            }
            else
                return style3;
            
        }
        private static DateTime baseDate = DateTime.Now;
        private static Random rnd = new Random();
        private static object GetRandomValue()
        {
            var val = rnd.Next(1, 7);
            if (val == 1)
                return Guid.NewGuid().ToString().Substring(1, 5);
            else if (val == 2)
                return true;
            else if (val == 3)
                return false;
            else if (val == 4)
                return DateTime.Now;
            else if (val == 5)
                return rnd.Next(1, 1000);
            else
                return (DateTime.Now - baseDate );
        }

        

        class Person
        {
            public String Name { get; set; }
            public Int32 Age { get; set; }
        }

        // Save defaults to a .config file
    }
}
