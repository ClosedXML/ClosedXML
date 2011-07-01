using System;
using System.Linq;
using ClosedXML.Excel;


namespace ClosedXML_Examples.Ranges
{
    public class CellMoves : IXLExample
    {
        #region Methods

        // Public
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Cell Moves");

            var cell = ws.Cell(5, 5).SetValue("(5,5)");

            cell.CellAbove().SetValue("(4,5)").Style.Fill.SetBackgroundColor(XLColor.LightSalmon);
            cell.CellAbove(2).SetValue("(3,5)").Style.Fill.SetBackgroundColor(XLColor.LightSalmon);
            cell.CellBelow().SetValue("(6,5)").Style.Fill.SetBackgroundColor(XLColor.Salmon);
            cell.CellBelow(2).SetValue("(7,5)").Style.Fill.SetBackgroundColor(XLColor.Salmon);

            cell.CellLeft().SetValue("(5,4)").Style.Fill.SetBackgroundColor(XLColor.LightBlue);
            cell.CellLeft(2).SetValue("(5,3)").Style.Fill.SetBackgroundColor(XLColor.LightBlue);
            cell.CellRight().SetValue("(5,6)").Style.Fill.SetBackgroundColor(XLColor.BlueBell);
            cell.CellRight(2).SetValue("(5,7)").Style.Fill.SetBackgroundColor(XLColor.BlueBell);

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
