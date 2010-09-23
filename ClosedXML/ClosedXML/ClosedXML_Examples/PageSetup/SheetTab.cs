using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;

namespace ClosedXML_Examples.PageSetup
{
    public class SheetTab
    {
        #region Methods

        // Public
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Sheet Tab");
            
            // Adding print areas
            ws.PageSetup.PrintAreas.Add(ws.Range("A1:B2"));
            ws.PageSetup.PrintAreas.Add(ws.Range("D3:D5"));

            // Adding rows to repeat at top
            ws.PageSetup.SetRowTitles(ws.Rows(1, 2));

            // Adding columns to repeat at left
            //ws.PageSetup.SetColumnTitles(ws.Columns(1, 2));

            // Show gridlines
            ws.PageSetup.ShowGridlines = true;

            // Print in black and white
            ws.PageSetup.BlackAndWhite = true;

            // Print in draft quality
            ws.PageSetup.DraftQuality = true;

            // Show row and column headings
            ws.PageSetup.ShowRowAndColumnHeadings = true;

            // Set the page print order to over, then down
            ws.PageSetup.PageOrder = XLPageOrderValues.OverThenDown;

            // Place comments at the end of the sheet
            ws.PageSetup.ShowComments = XLShowCommentsValues.AtEnd;

            // Print errors as #N/A
            ws.PageSetup.PrintErrorValue = XLPrintErrorValues.NA;
            
            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
