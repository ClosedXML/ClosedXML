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
            var ws = workbook.Worksheets.Add("Columns of a Range");

            // All columns in a range
            ws.Range("A1:B2").Columns().Style.Fill.BackgroundColor = Color.DimGray;

            var bigRange = ws.Range("A4:V6");

            // Contiguous columns by number
            bigRange.Columns(1, 2).Style.Fill.BackgroundColor = Color.Red;

            // Contiguous columns by letter
            bigRange.Columns("D", "E").Style.Fill.BackgroundColor = Color.Blue;

            // Contiguous columns by letter
            bigRange.Columns("G:H").Style.Fill.BackgroundColor = Color.DeepPink;

            // Spread columns by number
            bigRange.Columns("10:11,13:14").Style.Fill.BackgroundColor = Color.Orange;

            // Spread columns by letter
            bigRange.Columns("P:Q,S:T").Style.Fill.BackgroundColor = Color.Turquoise;

            // Use a single number/letter
            bigRange.Columns("V").Style.Fill.BackgroundColor = Color.Cyan;

            // Only the used columns in a worksheet
            ws.Columns("A:V").Width = 3; 


            var ws2 = workbook.Worksheets.Add("Columns of a worksheet");
            
            // Contiguous columns by number
            ws2.Columns(1, 2).Style.Fill.BackgroundColor = Color.Red;

            // Contiguous columns by letter
            ws2.Columns("D", "E").Style.Fill.BackgroundColor = Color.Blue;

            // Contiguous columns by letter
            ws2.Columns("G:H").Style.Fill.BackgroundColor = Color.DeepPink;

            // Spread columns by number
            ws2.Columns("10:11,13:14").Style.Fill.BackgroundColor = Color.Orange;

            // Spread columns by letter
            ws2.Columns("P:Q,S:T").Style.Fill.BackgroundColor = Color.Turquoise;

            // Use a single number/letter
            ws2.Columns("V").Style.Fill.BackgroundColor = Color.Cyan;

            ws2.Columns("A:V").Width = 3;
            
            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
