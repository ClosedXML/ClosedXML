using System;using System.Collections.Generic;
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


            FillStyles();
            var wb = new XLWorkbook();
            foreach (var i in Enumerable.Range(1, 1))
            {
                var ws = wb.Worksheets.Add("Sheet" + i);
                foreach (var ro in Enumerable.Range(1, 100))
                {
                    foreach (var co in Enumerable.Range(1, 100))
                    {
                        ws.Cell(ro, co).Style = GetRandomStyle();
                        ws.Cell(ro, co).Value = GetRandomValue();
                    }
                    System.Threading.Thread.Sleep(10);
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

            //var ws = wb.Worksheets.Worksheet(0);
            //wb.SaveAs(@"C:\Excel Files\ForTesting\Benchmark_Saved.xlsx");

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
            
            var val = rnd.Next(1, 3);
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
            var val = rnd.Next(1, 6);
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
