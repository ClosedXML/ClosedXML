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
            var ws = wb.Worksheets.Add("Print Areas");
            ws.PageSetup.PrintAreas.Add(ws.Range("A1:B2"));
            ws.PageSetup.PrintAreas.Add(ws.Range("D4:E5"));

            ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            ws.PageSetup.AdjustTo(85);
            var ws2 = wb.Worksheets.Add("Sheet2");
            ws2.PageSetup.PrintAreas.Add(ws2.Range("B2:E5"));
            ws2.PageSetup.PageOrientation = XLPageOrientation.Portrait;
            ws2.PageSetup.PagesWide = 1;
            ws2.PageSetup.PagesTall = 2;
            ws2.PageSetup.PaperSize = XLPaperSize.MonarchEnvelope;
            ws2.PageSetup.HorizontalDpi = 600;
            ws2.PageSetup.VerticalDpi = 600;
            ws2.PageSetup.FirstPageNumber = 6;
            ws2.PageSetup.CenterHorizontally = true;
            ws2.PageSetup.CenterVertically = true;
            ws2.PageSetup.Margins.Top = 1.5;

            var headerFont = new XLFont() { Bold = true };
            ws2.PageSetup.Header.Left.AddText("Test", XLHFOccurrence.OddPages, headerFont);
            ws2.PageSetup.Header.Left.AddText("Test", XLHFOccurrence.EvenPages, headerFont);
            ws2.PageSetup.Header.Left.AddText("Test", XLHFOccurrence.FirstPage, headerFont);
            ws2.PageSetup.Header.Left.AddText("Test", XLHFOccurrence.AllPages, headerFont);
            ws2.PageSetup.Header.Left.Clear();
            
            ws2.PageSetup.Footer.Center.AddText(XLHFPredefinedText.SheetName, XLHFOccurrence.AllPages, headerFont);
            ws2.PageSetup.DraftQuality = true;
            ws2.PageSetup.BlackAndWhite = true;
            ws2.PageSetup.PageOrder = XLPageOrderValues.OverThenDown;
            ws2.PageSetup.ShowComments = XLShowCommentsValues.AtEnd;
            ws2.PageSetup.PrintAreas.Add(ws2.Range("H10:H20"));
            ws2.PageSetup.RowTitles.Add(ws2.Row(1));

            
            
            // Add List<IXLRange> Ranges(...) to IXLRandge

            //Apply a style to the entire sheet (not just the used cells)

            wb.SaveAs(@"c:\Sandbox.xlsx");
            //Console.ReadKey();

        }
    }
}
