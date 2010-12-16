using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;

namespace ClosedXML_Examples.Misc
{
    public class MultipleSheets
    {
        #region Variables

        // Public

        // Private


        #endregion

        #region Properties

        // Public

        // Private

        // Override


        #endregion

        #region Events

        // Public

        // Private

        // Override


        #endregion

        #region Methods

        // Public
        public void Create()
        {
            var wb = new XLWorkbook(@"C:\Excel Files\Created\MultipleSheets.xlsx");
            var ws = wb.Worksheets.Add("NewOne");
            wb.Worksheets.Worksheet(0).Delete();
            ws.SheetIndex = 0;
            wb.Worksheets.Worksheet("Inserted").SheetIndex = wb.Worksheets.Count();
            wb.SaveAs(@"C:\Excel Files\Created\MultipleSheets_Saved.xlsx");

            wb = new XLWorkbook();
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

            wb.SaveAs(@"C:\Excel Files\Created\OrganizingSheets.xlsx");
        }

        // Private

        // Override


        #endregion
    }
}
