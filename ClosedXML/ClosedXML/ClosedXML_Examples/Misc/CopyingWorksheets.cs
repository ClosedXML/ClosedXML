using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;

namespace ClosedXML_Examples.Misc
{
    public class CopyingWorksheets
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
            var wb = new XLWorkbook(@"C:\Excel Files\Created\BasicTable.xlsx");
            var wsSource = wb.Worksheet(0);
            // Copy the worksheet to a new sheet in this workbook
            wsSource.CopyTo("Copy");

            // We're going to open another workbook to show that you can
            // copy a sheet from one workbook to another:
            var wbSource = new XLWorkbook(@"C:\Excel Files\Created\BasicTable.xlsx");
            wbSource.Worksheet(0).CopyTo(wb, "Copy From Other");

            // Save the workbook with the 2 copies
            wb.SaveAs(@"C:\Excel Files\Created\CopyingWorksheets.xlsx");
        }

        // Private

        // Override


        #endregion
    }
}
