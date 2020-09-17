using ClosedXML.Excel;

namespace ClosedXML_Examples.Misc
{
    public class DataValidationWholeNumber : IXLExample
    {
        public void Create(string filePath)
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var c1 = ws.Cell("A1");
            var c2 = ws.Cell("B1");
            c1.Value = 1;
            c2.Value = 2;

            ws.Range("A2:A10").SetDataValidation().WholeNumber.EqualTo(1);
            ws.Range("B2:B10").SetDataValidation().WholeNumber.NotEqualTo(2);
            ws.Range("C2:C10").SetDataValidation().WholeNumber.GreaterThan(3);
            ws.Range("D2:D10").SetDataValidation().WholeNumber.LessThan(4);
            ws.Range("E2:E10").SetDataValidation().WholeNumber.EqualOrGreaterThan(5);
            ws.Range("F2:F10").SetDataValidation().WholeNumber.EqualOrLessThan(6);
            ws.Range("G2:G10").SetDataValidation().WholeNumber.Between(7, 8);
            ws.Range("H2:H10").SetDataValidation().WholeNumber.NotBetween(9, 10);

            ws.Range("A11:A20").SetDataValidation().WholeNumber.EqualTo(c1);
            ws.Range("B11:B20").SetDataValidation().WholeNumber.NotEqualTo(c1);
            ws.Range("C11:C20").SetDataValidation().WholeNumber.GreaterThan(c1);
            ws.Range("D11:D20").SetDataValidation().WholeNumber.LessThan(c1);
            ws.Range("E11:E20").SetDataValidation().WholeNumber.EqualOrGreaterThan(c1);
            ws.Range("F11:F20").SetDataValidation().WholeNumber.EqualOrLessThan(c1);
            ws.Range("G11:G20").SetDataValidation().WholeNumber.Between(c1, c2);
            ws.Range("H11:H20").SetDataValidation().WholeNumber.NotBetween(c1, c2);

            wb.SaveAs(filePath);
        }
    }
}