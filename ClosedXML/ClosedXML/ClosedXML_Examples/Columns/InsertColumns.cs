using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.Excel.Style;
using System.Drawing;

namespace ClosedXML_Examples.Columns
{
    public class InsertColumns
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
            var ws = workbook.Worksheets.Add("Insert Columns");

            foreach (var r in Enumerable.Range(1, 5))
            {
                foreach (var c in Enumerable.Range(1, 5))
                {
                    ws.Cell(r, c).Value = "X";
                    ws.Cell(r, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }
            }

            ws.Column(3).InsertColumnsAfter(2);
            ws.Column(1).InsertColumnsBefore(2);
            ws.Range("D3:E4").InsertColumnsAfter(2);
            ws.Range("B3:C5").InsertColumnsBefore(2);

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
