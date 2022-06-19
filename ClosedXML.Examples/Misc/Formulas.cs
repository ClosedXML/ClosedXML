using ClosedXML.Excel;

namespace ClosedXML.Examples.Misc
{
    public class Formulas : IXLExample
    {
        public virtual void Create(string filePath)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Formulas");

            ws.Cell(1, 1).Value = "Num1";
            ws.Cell(1, 2).Value = "Num2";
            ws.Cell(1, 3).Value = "Total";
            ws.Cell(1, 4).Value = "cell.FormulaA1";
            ws.Cell(1, 5).Value = "cell.FormulaR1C1";
            ws.Cell(1, 6).Value = "cell.Value";
            ws.Cell(1, 7).Value = "Are Equal?";

            ws.Cell(2, 1).Value = 1;
            ws.Cell(2, 2).Value = 2;
            var cellWithFormulaA1 = ws.Cell(2, 3);
            // Use A1 notation
            cellWithFormulaA1.FormulaA1 = "=A2+$B$2"; // The equal sign (=) in a formula is optional
            ws.Cell(2, 4).Value = cellWithFormulaA1.FormulaA1;
            ws.Cell(2, 5).Value = cellWithFormulaA1.FormulaR1C1;
            ws.Cell(2, 6).Value = cellWithFormulaA1.Value;

            ws.Cell(3, 1).Value = 1;
            ws.Cell(3, 2).Value = 2;
            var cellWithFormulaR1C1 = ws.Cell(3, 3);
            // Use R1C1 notation
            cellWithFormulaR1C1.FormulaR1C1 = "RC[-2]+R3C2"; // The equal sign (=) in a formula is optional
            ws.Cell(3, 4).Value = cellWithFormulaR1C1.FormulaA1;
            ws.Cell(3, 5).Value = cellWithFormulaR1C1.FormulaR1C1;
            ws.Cell(3, 6).Value = cellWithFormulaR1C1.Value;

            ws.Cell(4, 1).Value = "A";
            ws.Cell(4, 2).Value = "B";
            var cellWithStringFormula = ws.Cell(4, 3);

            // Use R1C1 notation
            cellWithStringFormula.FormulaR1C1 = "=\"Test\" & RC[-2] & \"R3C2\"";
            ws.Cell(4, 4).Value = cellWithStringFormula.FormulaA1;
            ws.Cell(4, 5).Value = cellWithStringFormula.FormulaR1C1;
            ws.Cell(4, 6).Value = cellWithStringFormula.Value;

            // Setting the formula of a range
            var rngData = ws.Range(2, 1, 4, 7);
            rngData.LastColumn().FormulaR1C1 = "=IF(RC[-4]=RC[-1],\"Yes\", \"No\")";

            // Using an array formula:
            // Just put the formula between curly braces
            ws.Cell("A6").Value = "Array Formula: ";
            ws.Cell("B6").FormulaA1 = "{A2+A3}";
            ws.Range("C6:D6").FormulaA1 = "{TRANSPOSE(A2:A3)}";

            ws.Range(1, 1, 1, 7).Style.Fill.BackgroundColor = XLColor.Cyan;
            ws.Range(1, 1, 1, 7).Style.Font.Bold = true;
            ws.Columns().AdjustToContents();

            // You can also change the reference notation:
            wb.ReferenceStyle = XLReferenceStyle.R1C1;

            // And the workbook calculation mode:
            wb.CalculateMode = XLCalculateMode.Auto;

            ws.Range("A10").AddToNamed("A10_R1C1_A10_R1C1");
            ws.Cell("A10").Value = 0;
            ws.Cell("A11").FormulaA1 = "A2 + A10_R1C1_A10_R1C1";
            ws.Cell("A12").FormulaR1C1 = "R2C1 + A10_R1C1_A10_R1C1";
            ws.Cell("A13").FormulaR1C1 = "=SUM(R[-5]:R[-4])";
            ws.Cell("A14").FormulaA1 = "=SUM(8:9)";

            wb.SaveAs(filePath);
        }
    }
}