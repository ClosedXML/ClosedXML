using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.Excel.Style;
using System.Drawing;

namespace ClosedXML_Examples.Columns
{
    public class ColumnSettings
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
        public ColumnSettings()
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
            var ws = workbook.Worksheets.Add("Column Settings");

            ws.Cell("B4").Style.Fill.BackgroundColor = Color.Brown;
            ws.Column("B").Style.Fill.BackgroundColor = Color.Red;
            ws.Cell("B2").Style.Fill.BackgroundColor = Color.Blue;
            ws.Column("B").Width = 15;
            ws.Column(4).Style.Fill.BackgroundColor = Color.DarkOrange;

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
