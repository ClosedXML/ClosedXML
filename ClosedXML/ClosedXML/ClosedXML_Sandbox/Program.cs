using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.Excel.Style;
using System.Drawing;

namespace ClosedXML_Sandbox
{
    class Program
    {
        static void Main(string[] args)
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            ws1.PrintOptions.PrintArea = ws1.Range("A1:B2");
            ws1.PrintOptions.PageOrientation = XLPageOrientation.Portrait;
            var ws2 = wb.Worksheets.Add("Sheet2");
            ws2.PrintOptions.PrintArea = ws2.Range("B2:E5");
            ws2.PrintOptions.PageOrientation = XLPageOrientation.Landscape;
            ws2.PrintOptions.PagesWide = 1;
            ws2.PrintOptions.PagesTall = 2;
            foreach (var wsNum in Enumerable.Range(1, 5))
            {
                var ws = wb.Worksheets.Add("New Sheet " + wsNum.ToString());
                foreach (var ro in Enumerable.Range(1, 5))
                {
                    foreach (var co in Enumerable.Range(1, 5))
                    {
                        ws.Cell(ro, co).Value = "(" + ro + "," + co + ")";
                    }
                }
            }


            //foreach (var c in ws.Range("B2:C3").Columns())
            //{
            //    c.Style.Fill.BackgroundColor = Color.Red;
            //}

            //ws.Cell("E1").Value = "Wide 2";

            //foreach (var c in ws.Columns())
            //{
            //    c.Width = 20;
            //}

            // Fix Worksheet.Cells() method   !!!




            //foreach (var r in ws.Rows())
            //{
            //    r.Height = 15;
            //}

            //foreach (var r in ws.Range("B2:C3").Rows())
            //{
            //    r.Style.Fill.BackgroundColor = Color.Red;
            //}

            //ws.Columns("A:B").Width = 20;
            //ws.Columns("3:4").Width = 20;
            //ws.Rows("1:2").Height = 30;

            //ws.Columns("A:B").Style.Fill.BackgroundColor = Color.Red;
            //ws.Columns("3:4").Style.Fill.BackgroundColor = Color.Blue;
            //ws.Rows("1:2").Style.Fill.BackgroundColor = Color.Orange;

            //var rng1 = ws.Range("B2:E5");
            //rng1.Columns("A:B").Style.Fill.BackgroundColor = Color.Red;
            //rng1.Columns("3:4").Style.Fill.BackgroundColor = Color.Blue;
            //rng1.Rows("1:2").Style.Fill.BackgroundColor = Color.Orange;

            //ws.Row(2).Delete();
            //ws.Column(2).Delete();
            //ws.Column("B").Delete();

            //ws.Columns("A:B").Delete();
            //ws.Columns("3:4").Delete();
            //ws.Rows("1:2").Delete();

            //ws.Range("B2:C3").Delete(ShiftCellsUp);
            //ws.Range("B2:C3").Delete(ShiftCellsLeft);

            //ws.Range("B2:C3").Column(1).Delete(ShiftCellsUp);
            //ws.Range("B2:C3").Column("A").Delete(ShiftCellsLeft);
            //ws.Range("B2:C3").Row(1).Delete(ShiftCellsUp);
            //ws.Range("B2:C3").Row((1).Delete(ShiftCellsLeft);

            wb.SaveAs(@"c:\Sandbox.xlsx");
            //Console.ReadKey();

        }
    }
}
