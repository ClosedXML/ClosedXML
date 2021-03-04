using ClosedXML.Excel;

namespace ClosedXML.Examples.Misc
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

            ws.Range("A2:A10").CreateDataValidation().TextLength.EqualTo(1);
            ws.Range("B2:B10").CreateDataValidation().TextLength.NotEqualTo(2);
            ws.Range("C2:C10").CreateDataValidation().TextLength.GreaterThan(3);
            ws.Range("D2:D10").CreateDataValidation().TextLength.LessThan(4);
            ws.Range("E2:E10").CreateDataValidation().TextLength.EqualOrGreaterThan(5);
            ws.Range("F2:F10").CreateDataValidation().TextLength.EqualOrLessThan(6);
            ws.Range("G2:G10").CreateDataValidation().TextLength.Between(7, 8);
            ws.Range("H2:H10").CreateDataValidation().TextLength.NotBetween(9, 10);

            ws.Range("A11:A20").CreateDataValidation().TextLength.EqualTo(c1);
            ws.Range("B11:B20").CreateDataValidation().TextLength.NotEqualTo(c1);
            ws.Range("C11:C20").CreateDataValidation().TextLength.GreaterThan(c1);
            ws.Range("D11:D20").CreateDataValidation().TextLength.LessThan(c1);
            ws.Range("E11:E20").CreateDataValidation().TextLength.EqualOrGreaterThan(c1);
            ws.Range("F11:F20").CreateDataValidation().TextLength.EqualOrLessThan(c1);
            ws.Range("G11:G20").CreateDataValidation().TextLength.Between(c1, c2);
            ws.Range("H11:H20").CreateDataValidation().TextLength.NotBetween(c1, c2);

            wb.SaveAs(filePath);
        }
    }
}
