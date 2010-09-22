using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;

namespace ClosedXML_Examples.Rows
{
    public class RowCollection
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
            var ws = workbook.Worksheets.Add("Row Collection");

            // All rows in a range
            ws.Range("B2:C3").Rows().ForEach(r => r.Style.Fill.BackgroundColor = Color.Red);

            // Let's add a separate cell to the worksheet
            ws.Cell("B5").Value = "Tall 2";

            // Only the used rows in a worksheet
            ws.Rows().ForEach(r => r.Height = 30);

            var ws2 = workbook.Worksheets.Add("Multiple Rows");

            // Contiguous rows by number
            ws2.Rows(1, 2).ForEach(r => r.Style.Fill.BackgroundColor = Color.Red);

            // Contiguous rows by number
            ws2.Rows("4:5").ForEach(r => r.Style.Fill.BackgroundColor = Color.Blue);

            // Spread rows by number
            ws2.Rows("7:8,10:11").ForEach(r => r.Style.Fill.BackgroundColor = Color.Orange);

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
