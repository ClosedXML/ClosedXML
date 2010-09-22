using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

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

            // All columns in a range
            ws.Range("B2:C3").Columns().ForEach(c => c.Style.Fill.BackgroundColor = Color.Red);

            // Let's add a separate cell to the worksheet
            ws.Cell("E1").Value = "Wide 2";

            // Only the used columns in a worksheet
            ws.Columns().ForEach(c => c.Width = 20);

            var ws2 = workbook.Worksheets.Add("Multiple Columns");
            
            // Contiguous columns by number
            ws2.Columns(1, 2).ForEach(c => c.Style.Fill.BackgroundColor = Color.Red);

            // Contiguous columns by letter
            ws2.Columns("D", "E").ForEach(c => c.Style.Fill.BackgroundColor = Color.Blue);

            // Contiguous columns by letter
            ws2.Columns("G:H").ForEach(c => c.Style.Fill.BackgroundColor = Color.Blue);

            // Spread columns by number
            ws2.Columns("10:11,13:14").ForEach(c => c.Style.Fill.BackgroundColor = Color.Orange);

            // Spread columns by letter
            ws2.Columns("P:Q,S:T").ForEach(c => c.Style.Fill.BackgroundColor = Color.Turquoise);

            ws2.Columns("A:T").ForEach(c => c.Width = 3);

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
