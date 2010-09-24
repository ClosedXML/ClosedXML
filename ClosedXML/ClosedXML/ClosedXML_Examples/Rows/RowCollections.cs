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
            var ws = workbook.Worksheets.Add("Rows of a Range");

            // All rows in a range
            ws.Range("A1:B2").Rows().Style.Fill.BackgroundColor = Color.DimGray;

            var bigRange = ws.Range("B4:C17");

            // Contiguous rows by number
            bigRange.Rows(1, 2).Style.Fill.BackgroundColor = Color.Red;

            // Contiguous rows by number
            bigRange.Rows("4:5").Style.Fill.BackgroundColor = Color.Blue;

            // Spread rows by number
            bigRange.Rows("7:8,10:11").Style.Fill.BackgroundColor = Color.Orange;

            // Using a single number
            bigRange.Rows("13").Style.Fill.BackgroundColor = Color.Cyan;

            // Only the used rows in a worksheet
            ws.Rows().Height = 15;

            var ws2 = workbook.Worksheets.Add("Rows of a Worksheet");

            // Contiguous rows by number
            ws2.Rows(1, 2).Style.Fill.BackgroundColor = Color.Red;

            // Contiguous rows by number
            ws2.Rows("4:5").Style.Fill.BackgroundColor = Color.Blue;

            // Spread rows by number
            ws2.Rows("7:8,10:11").Style.Fill.BackgroundColor = Color.Orange;

            // Using a single number
            ws2.Rows("13").Style.Fill.BackgroundColor = Color.Cyan;

            ws2.Rows("1:13").Height = 15;

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
