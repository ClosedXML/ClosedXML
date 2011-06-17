using System;
using ClosedXML.Excel;


namespace ClosedXML_Examples
{
    public class ColumnCells
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
            var ws = workbook.Worksheets.Add("Column Cells");

            var columnFromWorksheet = ws.Column(1);
            columnFromWorksheet.Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
            columnFromWorksheet.Cells("2").Style.Fill.BackgroundColor = XLColor.Blue;
            columnFromWorksheet.Cells("3,5:6").Style.Fill.BackgroundColor = XLColor.Red;
            columnFromWorksheet.Cells(8, 9).Style.Fill.BackgroundColor = XLColor.Blue;

            var columnFromRange = ws.Range("B1:B9").FirstColumn();

            columnFromRange.Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
            columnFromRange.Cells("2").Style.Fill.BackgroundColor = XLColor.Blue;
            columnFromRange.Cells("3,5:6").Style.Fill.BackgroundColor = XLColor.Red;
            columnFromRange.Cells(8, 9).Style.Fill.BackgroundColor = XLColor.Blue;

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
