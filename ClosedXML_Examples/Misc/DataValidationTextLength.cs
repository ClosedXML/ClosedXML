using ClosedXML.Excel;

namespace ClosedXML_Examples.Misc
{
    public class DataValidationTextLength : IXLExample
    {
        public void Create(string filePath)
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var c1 = ws.Cell("A1");
            var c2 = ws.Cell("B1");
            c1.Value = 1;
            c2.Value = 2;

            ws.Range("A2:A10").SetDataValidation().TextLength.EqualTo(1);
            ws.Range("B2:B10").SetDataValidation().TextLength.NotEqualTo(2);
            ws.Range("C2:C10").SetDataValidation().TextLength.GreaterThan(3);
            ws.Range("D2:D10").SetDataValidation().TextLength.LessThan(4);
            ws.Range("E2:E10").SetDataValidation().TextLength.EqualOrGreaterThan(5);
            ws.Range("F2:F10").SetDataValidation().TextLength.EqualOrLessThan(6);
            ws.Range("G2:G10").SetDataValidation().TextLength.Between(7, 8);
            ws.Range("H2:H10").SetDataValidation().TextLength.NotBetween(9, 10);

            ws.Range("A11:A20").SetDataValidation().TextLength.EqualTo(c1);
            ws.Range("B11:B20").SetDataValidation().TextLength.NotEqualTo(c1);
            ws.Range("C11:C20").SetDataValidation().TextLength.GreaterThan(c1);
            ws.Range("D11:D20").SetDataValidation().TextLength.LessThan(c1);
            ws.Range("E11:E20").SetDataValidation().TextLength.EqualOrGreaterThan(c1);
            ws.Range("F11:F20").SetDataValidation().TextLength.EqualOrLessThan(c1);
            ws.Range("G11:G20").SetDataValidation().TextLength.Between(c1, c2);
            ws.Range("H11:H20").SetDataValidation().TextLength.NotBetween(c1, c2);

            wb.SaveAs(filePath);
        }
    }
}