using ClosedXML.Excel;

namespace ClosedXML_Examples.Misc
{
    public class DataValidationDecimal : IXLExample
    {
        public void Create(string filePath)
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var c1 = ws.Cell("A1");
            var c2 = ws.Cell("B1");
            c1.Value = 1.1;
            c2.Value = 2.1;
            var r1 = ws.Range("A1:A10");
            var r2 = ws.Range("B1:B10");

            ws.Range("A2:A10").SetDataValidation().Decimal.EqualTo(1.1);
            ws.Range("B2:B10").SetDataValidation().Decimal.NotEqualTo(2.1);
            ws.Range("C2:C10").SetDataValidation().Decimal.GreaterThan(3.1);
            ws.Range("D2:D10").SetDataValidation().Decimal.LessThan(4.1);
            ws.Range("E2:E10").SetDataValidation().Decimal.EqualOrGreaterThan(5.1);
            ws.Range("F2:F10").SetDataValidation().Decimal.EqualOrLessThan(6.1);
            ws.Range("G2:G10").SetDataValidation().Decimal.Between(7.1, 8.1);
            ws.Range("H2:H10").SetDataValidation().Decimal.NotBetween(9.1, 10.1);

            ws.Range("A11:A20").SetDataValidation().Decimal.EqualTo(c1);
            ws.Range("B11:B20").SetDataValidation().Decimal.NotEqualTo(c1);
            ws.Range("C11:C20").SetDataValidation().Decimal.GreaterThan(c1);
            ws.Range("D11:D20").SetDataValidation().Decimal.LessThan(c1);
            ws.Range("E11:E20").SetDataValidation().Decimal.EqualOrGreaterThan(c1);
            ws.Range("F11:F20").SetDataValidation().Decimal.EqualOrLessThan(c1);
            ws.Range("G11:G20").SetDataValidation().Decimal.Between(c1, c2);
            ws.Range("H11:H20").SetDataValidation().Decimal.NotBetween(c1, c2);

            wb.SaveAs(filePath);
        }
    }
}