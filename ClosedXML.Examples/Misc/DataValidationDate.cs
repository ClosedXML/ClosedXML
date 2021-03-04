using ClosedXML.Excel;
using System;

namespace ClosedXML.Examples.Misc
{
    public class DataValidationDate : IXLExample
    {
        public void Create(string filePath)
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var date1 = new DateTime(2020, 01, 31);
            var date2 = new DateTime(2020, 02, 29);
            var c1 = ws.Cell("A1");
            var c2 = ws.Cell("B1");
            c1.Value = date1;
            c2.Value = date2;

            ws.Range("A2:A10").CreateDataValidation().Date.EqualTo(date1);
            ws.Range("B2:B10").CreateDataValidation().Date.NotEqualTo(date1);
            ws.Range("C2:C10").CreateDataValidation().Date.GreaterThan(date1);
            ws.Range("D2:D10").CreateDataValidation().Date.LessThan(date1);
            ws.Range("E2:E10").CreateDataValidation().Date.EqualOrGreaterThan(date1);
            ws.Range("F2:F10").CreateDataValidation().Date.EqualOrLessThan(date1);
            ws.Range("G2:G10").CreateDataValidation().Date.Between(date1, date2);
            ws.Range("H2:H10").CreateDataValidation().Date.NotBetween(date1, date2);

            ws.Range("A11:A20").CreateDataValidation().Date.EqualTo(c1);
            ws.Range("B11:B20").CreateDataValidation().Date.NotEqualTo(c1);
            ws.Range("C11:C20").CreateDataValidation().Date.GreaterThan(c1);
            ws.Range("D11:D20").CreateDataValidation().Date.LessThan(c1);
            ws.Range("E11:E20").CreateDataValidation().Date.EqualOrGreaterThan(c1);
            ws.Range("F11:F20").CreateDataValidation().Date.EqualOrLessThan(c1);
            ws.Range("G11:G20").CreateDataValidation().Date.Between(c1, c2);
            ws.Range("H11:H20").CreateDataValidation().Date.NotBetween(c1, c2);

            wb.SaveAs(filePath);
        }
    }
}
