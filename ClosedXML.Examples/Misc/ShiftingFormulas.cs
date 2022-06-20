using ClosedXML.Excel;
using System.Linq;

namespace ClosedXML.Examples.Misc
{
    public class ShiftingFormulas : IXLExample
    {
        #region Variables

        // Public

        // Private

        #endregion Variables

        #region Properties

        // Public

        // Private

        // Override

        #endregion Properties

        #region Events

        // Public

        // Private

        // Override

        #endregion Events

        #region Methods

        // Public
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Shifting Formulas");
            ws.Cell("B2").Value = 5;
            ws.Cell("B3").Value = 6;
            ws.Cell("C2").Value = 1;
            ws.Cell("C3").Value = 2;
            ws.Cell("A4").Value = "Sum:";
            ws.Range("B4:C4").FormulaR1C1 = "Sum(R[-2]C:R[-1]C)";
            ws.Range("B4:C4").AddToNamed("WorkbookB4C4");
            ws.Range("B4:C4").AddToNamed("WorksheetB4C4", XLScope.Worksheet);
            ws.Cell("E2").Value = "Avg:";

            ws.Cell("F2").FormulaA1 = "Average(B2:C3)";
            ws.Ranges("A4,E2").Style
                .Font.SetBold()
                .Fill.SetBackgroundColor(XLColor.CyanProcess);

            var ws2 = wb.Worksheets.Add("WS2");
            ws2.Cell(1, 1).FormulaA1 = "='Shifting Formulas'!B2";
            ws2.Cell(1, 2).Value = ws2.Cell(1, 1).Value;
            ws2.Cell(2, 1).FormulaA1 = "Average('Shifting Formulas'!$B$2:$C$3)";
            ws2.Cell(3, 1).FormulaA1 = "Average('Shifting Formulas'!$B$2:$C3)";
            ws2.Cell(4, 1).FormulaA1 = "Average('Shifting Formulas'!$B$2:C3)";
            ws2.Cell(5, 1).FormulaA1 = "Average('Shifting Formulas'!$B2:C3)";
            ws2.Cell(6, 1).FormulaA1 = "Average('Shifting Formulas'!B2:C3)";
            ws2.Cell(7, 1).FormulaA1 = "Average('Shifting Formulas'!B2:C$3)";
            ws2.Cell(8, 1).FormulaA1 = "Average('Shifting Formulas'!B2:$C$3)";
            ws2.Cell(9, 1).FormulaA1 = "Average('Shifting Formulas'!B$2:$C$3)";

            var dataGrid = ws.Range("B2:D3");
            ws.Row(1).InsertRowsAbove(1);
            var newRow = dataGrid.LastRow().InsertRowsAbove(1).First();
            newRow.Value = 1;
            dataGrid.LastColumn().FormulaR1C1 = string.Format("SUM(RC[-{0}]:RC[-1])", dataGrid.ColumnCount() - 1);
            ws.Cell(1, 1).InsertCellsBelow(1);
            ws.Column(1).InsertColumnsBefore(1);
            ws.Row(4).Delete();
            wb.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}