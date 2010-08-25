using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.Excel.Style;
using System.Drawing;

namespace ClosedXML_Examples.Rows
{
    public class RowSettings
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

        #region Constructors

        // Public
        public RowSettings()
        {

        }


        // Private


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
            var ws = workbook.Worksheets.Add("Row Settings");

            ws.Cell("D2").Style.Fill.BackgroundColor = Color.Brown;
            ws.Row(2).Style.Fill.BackgroundColor = Color.Red;
            ws.Cell("B2").Style.Fill.BackgroundColor = Color.Blue;
            ws.Row(2).Height = 30;
            ws.Row(4).Style.Fill.BackgroundColor = Color.DarkOrange;

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
