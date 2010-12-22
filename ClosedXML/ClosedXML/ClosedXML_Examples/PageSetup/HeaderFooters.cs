using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;

namespace ClosedXML_Examples.PageSetup
{
    public class HeaderFooters
    {
        #region Methods

        // Public
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Headers and Footers");
            
            // Simple left header to be placed on all pages
            ws.PageSetup.Header.Left.AddText("Created with ClosedXML");

            // Using various fonts for the right header on the first page only
            var font1 = XLWorkbook.GetXLFont();
            font1.Bold = true;
            ws.PageSetup.Header.Right.AddText("The ", XLHFOccurrence.FirstPage, font1);

            var font2 = XLWorkbook.GetXLFont();
            font2.FontColor = XLColor.Red;
            ws.PageSetup.Header.Right.AddText("First ", XLHFOccurrence.FirstPage, font2);

            var font3 = XLWorkbook.GetXLFont();
            font3.Underline = XLFontUnderlineValues.Double;
            ws.PageSetup.Header.Right.AddText("Colorful ", XLHFOccurrence.FirstPage, font3);

            var font4 = XLWorkbook.GetXLFont();
            font4.FontName = "Broadway";
            ws.PageSetup.Header.Right.AddText("Page", XLHFOccurrence.FirstPage, font4);

            // Using predefined header/footer text:

            // Let's put the full path to the file on the right footer of every odd page:
            ws.PageSetup.Footer.Right.AddText(XLHFPredefinedText.FullPath, XLHFOccurrence.OddPages);

            // Let's put the current page number and total pages on the center of every footer:
            ws.PageSetup.Footer.Center.AddText(XLHFPredefinedText.PageNumber, XLHFOccurrence.AllPages);
            ws.PageSetup.Footer.Center.AddText(" / ", XLHFOccurrence.AllPages);
            ws.PageSetup.Footer.Center.AddText(XLHFPredefinedText.NumberOfPages, XLHFOccurrence.AllPages);

            // Don't align headers and footers with the margins
            ws.PageSetup.AlignHFWithMargins = false;

            // Don't scale headers and footers with the document
            ws.PageSetup.ScaleHFWithDocument = false;

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
