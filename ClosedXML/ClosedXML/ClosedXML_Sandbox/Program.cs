using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ClosedXML;
using System.Drawing;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace ClosedXML_Sandbox
{
    class Program
    {
        static void xMain(string[] args)
        {
            //var fileName = "DataValidation";
            var fileName = "Sandbox";
            //var fileName = "Issue_0000";
            //var wb = new XLWorkbook(String.Format(@"c:\Excel Files\ForTesting\{0}.xlsx", fileName));
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("SheetX");
            ws.Cell(1, 1).Value = "1 234";
            
            //ws.Cell("A1").Value = "Category";
            //ws.Cell("A2").Value = "A";
            //ws.Cell("A3").Value = "B";
            //ws.Cell("B1").Value = "Value";
            //ws.Cell("B2").Value = 5;
            //ws.Cell("B3").Value = 10;

            //ws.RangeUsed().CreateChart(4, 4, 22, 12);

            //workbook.Worksheet("CCR").Column("C").LastCellUsed().Address is: {C29}
            //workbook.Worksheet("CCR").Column("B").LastCellUsed().Address is : {B25}

            //Now, when i use workbook.Worksheet("CCR").Range("B1:C34").RangeUsed().
            //The expect is B1:C29.

            wb.SaveAs(String.Format(@"c:\Excel Files\ForTesting\{0}_Saved.xlsx", fileName));


            //var start = DateTime.Now;
            //var wb = new XLWorkbook(@"C:\Excel Files\ForTesting\Issue_0000.xlsx");
            //var end = DateTime.Now;
            //Console.WriteLine(String.Format("Opened file in {0} seconds", (end - start).TotalSeconds));
            //var ws = wb.Worksheet(1);
            //var cell = ws.Cell(100000, 13);
            //Console.WriteLine(cell.GetString());
            Console.ReadKey();
        }



        static void CopyWorksheets(String source, XLWorkbook target)
        {
            var wb = new XLWorkbook(source);
            foreach (var ws in wb.Worksheets)
            {
                ws.CopyTo(target, ws.Name);
            }
        }

        static void Main(string[] args)
        {
            FillStyles();
            List<Double> runningSave = new List<Double>();
            List<Double> runningLoad = new List<Double>();
            List<Double> runningSavedBack = new List<Double>();

            var wb = new XLWorkbook();
            var startTotal = DateTime.Now;
            var start = DateTime.Now;
            foreach (var i in Enumerable.Range(1, 1))
            {
                var ws = wb.Worksheets.Add("Sheet" + i);
                foreach (var ro in Enumerable.Range(1, 10000))
                {
                    foreach (var co in Enumerable.Range(1, 20))
                    {
                        //ws.Cell(ro, co).Style = GetRandomStyle();
                        //if (rnd.Next(1, 5) == 1)
                        //ws.Cell(ro, co).FormulaA1 = ws.Cell(ro + 1, co + 1).Address.ToString() + " & \"-Copy\"";
                        //else
                        ws.Cell(ro, co).Value = GetRandomValue();
                    }
                    //System.Threading.Thread.Sleep(10);
                }
            }

            var end = DateTime.Now;
            Console.WriteLine(String.Format("Created file in {0} seconds", (end - start).TotalSeconds));

            start = DateTime.Now;
            wb.SaveAs(@"C:\Excel Files\ForTesting\Benchmark.xlsx");
            end = DateTime.Now;
            var saved = (end - start).TotalSeconds;
            runningSave.Add(saved);
            Console.WriteLine("Saved in {0} secs.", saved);

            foreach (Int32 r in Enumerable.Range(1, 1))
            {
                var start1 = DateTime.Now;
                var wb1 = new XLWorkbook(@"C:\Excel Files\ForTesting\Benchmark.xlsx");
                var end1 = DateTime.Now;
                var loaded = (end1 - start1).TotalSeconds;
                runningLoad.Add(loaded);
                Console.WriteLine("Loaded in {0} secs.", loaded);

                //var start2 = DateTime.Now;
                ////wb1.SaveAs(@"C:\Excel Files\ForTesting\Benchmark_Saved.xlsx");
                //var end2 = DateTime.Now;
                //var savedBack = (end2 - start2).TotalSeconds;
                //runningSavedBack.Add(savedBack);
                //Console.WriteLine("Saved back in {0} secs.", savedBack);

                var endTotal = DateTime.Now;
                Console.WriteLine("It all took {0} secs.", (endTotal - startTotal).TotalSeconds);
            }
            Console.WriteLine("-------");
            Console.WriteLine("Avg Save time: {0}", runningSave.Average());
            Console.WriteLine("Avg Load time: {0}", runningLoad.Average());
            //Console.WriteLine("Avg Save Back time: {0}", runningSavedBack.Average());
            Console.ReadKey();
        }

        private static IXLStyle style1;
        private static IXLStyle style2;
        private static IXLStyle style3;
        private static void FillStyles()
        {

            style1 = XLWorkbook.DefaultStyle;
            style1.Font.Bold = true;
            style1.Fill.BackgroundColor = XLColor.Azure;
            style1.Border.BottomBorder = XLBorderStyleValues.Medium;
            style1.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            style2 = XLWorkbook.DefaultStyle;
            style2.Font.Italic = true;
            style2.Fill.BackgroundColor = XLColor.Orange;
            style2.Border.LeftBorder = XLBorderStyleValues.Medium;
            style2.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            style3 = XLWorkbook.DefaultStyle;
            style3.Font.FontColor = XLColor.FromColor(Color.Red);
            style3.Fill.PatternColor = XLColor.Blue;
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
                return (DateTime.Now - baseDate);
        }



        class Person
        {
            public String Name { get; set; }
            public Int32 Age { get; set; }
        }

        // Save defaults to a .config file
    }
}
