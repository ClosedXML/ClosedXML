using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.Excel.Style;
using System.Drawing;

namespace ClosedXML_Examples.Columns
{
    public class ColumnCollection
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
            var ws = workbook.Worksheets.Add("Column Collection");

            foreach (var c in ws.Range("B2:C3").Columns())
            {
                c.Style.Fill.BackgroundColor = Color.Red;
            }

            ws.Cell("E1").Value = "Wide 2";

            foreach (var c in ws.Columns())
            {
                c.Width = 20;
            }

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
