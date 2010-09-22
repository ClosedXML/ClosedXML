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
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            foreach (var wsNum in Enumerable.Range(1, 5))
            {
                var ws = workbook.Worksheets.Add("Sheet " + wsNum.ToString());
            }

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
