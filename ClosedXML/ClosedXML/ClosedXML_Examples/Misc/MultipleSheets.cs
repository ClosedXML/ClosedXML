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
        }

        // Private

        // Override


        #endregion
    }
}
