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
        static void Main(string[] args)
        {
            //var fileName = "DataValidation";
            var fileName = "Sandbox";
            //var fileName = "Issue_6724";
            //var wb = new XLWorkbook(String.Format(@"c:\Excel Files\ForTesting\{0}.xlsx", fileName));
            var wb = new XLWorkbook();
            //var ws = wb.Worksheets.Add("Sheet1");
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

            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cells("A1:C1").Value = "Initial";
            ws.Cells("C29,B25").Value = "Final";
            

            Console.WriteLine(ws.Column("C").LastCellUsed().Address.ToString( XLReferenceStyle.A1) );
            Console.WriteLine(ws.Column("B").LastCellUsed().Address.ToString(XLReferenceStyle.R1C1 ));
            Console.WriteLine(ws.Range("B1:C34").RangeUsed().RangeAddress.ToString());
            //wb.SaveAs(String.Format(@"c:\Excel Files\ForTesting\{0}_Saved.xlsx", fileName));
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

        static void xMain(string[] args)
        {
            FillStyles();
            List<Double> runningSave = new List<Double>();
            List<Double> runningLoad = new List<Double>();
            List<Double> runningSavedBack = new List<Double>();
            
            foreach (Int32 r in Enumerable.Range(1, 1))
            {
                var wb = new XLWorkbook();
                var startTotal = DateTime.Now;
                foreach (var i in Enumerable.Range(1, 1))
                {
                    var ws = wb.Worksheets.Add("Sheet" + i);
                    foreach (var ro in Enumerable.Range(1, 2000))
                    {
                        foreach (var co in Enumerable.Range(1, 100))
                        {
                            ws.Cell(ro, co).Style = GetRandomStyle();
                            //if (rnd.Next(1, 5) == 1)
                                //ws.Cell(ro, co).FormulaA1 = ws.Cell(ro + 1, co + 1).Address.ToString() + " & \"-Copy\"";
                            //else
                                ws.Cell(ro, co).Value = GetRandomValue();
                        }
                        //System.Threading.Thread.Sleep(10);
                    }
                    
                    //Int32 rowCount = ws.LastRowUsed().RowNumber();
                    //for (Int32 ro = 1; ro <= rowCount; ro += 100)
                    //{
                    //    var dv = ws.Range(ro, 1, ro + 99, 5).DataValidation;
                    //}

                    //var rngUsed = ws.RangeUsed();
                    //ws.RangeUsed().Style.Border.BottomBorder = XLBorderStyleValues.DashDot;
                    //ws.RangeUsed().Style.Border.BottomBorderColor = XLColor.AirForceBlue;
                    //ws.RangeUsed().Style.Border.TopBorder = XLBorderStyleValues.DashDotDot;
                    //ws.RangeUsed().Style.Border.TopBorderColor = XLColor.AliceBlue;
                    //ws.RangeUsed().Style.Border.LeftBorder = XLBorderStyleValues.Dashed;
                    //ws.RangeUsed().Style.Border.LeftBorderColor = XLColor.Alizarin;
                    //ws.RangeUsed().Style.Border.RightBorder = XLBorderStyleValues.Dotted;
                    //ws.RangeUsed().Style.Border.RightBorderColor = XLColor.Almond;

                    //ws.RangeUsed().Style.Font.Bold = true;
                    //ws.RangeUsed().Style.Font.FontColor = XLColor.Amaranth;
                    //ws.RangeUsed().Style.Font.FontSize = 10;
                    //ws.RangeUsed().Style.Font.Italic = true;

                    //ws.RangeUsed().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    //ws.RangeUsed().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    //ws.RangeUsed().Style.Alignment.WrapText = true;

                    //var startS = DateTime.Now;
                    //ws.Sort();
                    //var endS = DateTime.Now;
                    //var savedS = (endS - startS).TotalSeconds;
                    //runningSave.Add(savedS);
                    //Console.WriteLine("Sorted in {0} secs.", savedS);
                }


                //var start3 = DateTime.Now;
                //foreach (var ws in wb.Worksheets)
                //{
                //    ws.Style = wb.Style;
                //}
                //var end3 = DateTime.Now;
                //Console.WriteLine("Bolded all cells in {0} secs.", (end3 - start3).TotalSeconds);
                
                var start = DateTime.Now;
                wb.SaveAs(@"C:\Excel Files\ForTesting\Benchmark.xlsx");
                var end = DateTime.Now;
                var saved = (end - start).TotalSeconds;
                runningSave.Add(saved);
                Console.WriteLine("Saved in {0} secs.", saved);

                var start1 = DateTime.Now;
                var wb1 = new XLWorkbook(@"C:\Excel Files\ForTesting\Benchmark.xlsx");
                var end1 = DateTime.Now;
                var loaded = (end1 - start1).TotalSeconds;
                runningLoad.Add(loaded);
                Console.WriteLine("Loaded in {0} secs.", loaded);

                var start2 = DateTime.Now;
                wb1.SaveAs(@"C:\Excel Files\ForTesting\Benchmark_Saved.xlsx");
                var end2 = DateTime.Now;
                var savedBack = (end2 - start2).TotalSeconds;
                runningSavedBack.Add(savedBack);
                Console.WriteLine("Saved back in {0} secs.", savedBack);

                var endTotal = DateTime.Now;
                Console.WriteLine("It all took {0} secs.", (endTotal - startTotal).TotalSeconds);
            }
            Console.WriteLine("-------");
            Console.WriteLine("Avg Save time: {0}", runningSave.Average());
            Console.WriteLine("Avg Load time: {0}", runningLoad.Average());
            Console.WriteLine("Avg Save Back time: {0}", runningSavedBack.Average());
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
