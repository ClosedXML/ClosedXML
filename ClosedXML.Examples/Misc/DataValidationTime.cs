using ClosedXML.Excel;
using System;

namespace ClosedXML.Examples.Misc
{
    public class DataValidationTime : IXLExample
    {
        public void Create(string filePath)
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var time1 = TimeSpan.FromHours(6);
            var time2 = TimeSpan.FromHours(12);
            var c1 = ws.Cell("A1");
            var c2 = ws.Cell("B1");
            c1.Value = time1;
            c2.Value = time2;

            ws.Range("A2:A10").CreateDataValidation().Time.EqualTo(time1);
            ws.Range("B2:B10").CreateDataValidation().Time.NotEqualTo(time1);
            ws.Range("C2:C10").CreateDataValidation().Time.GreaterThan(time1);
            ws.Range("D2:D10").CreateDataValidation().Time.LessThan(time1);
            ws.Range("E2:E10").CreateDataValidation().Time.EqualOrGreaterThan(time1);
            ws.Range("F2:F10").CreateDataValidation().Time.EqualOrLessThan(time1);
            ws.Range("G2:G10").CreateDataValidation().Time.Between(time1, time2);
            ws.Range("H2:H10").CreateDataValidation().Time.NotBetween(time1, time2);

            ws.Range("A11:A20").CreateDataValidation().Time.EqualTo(c1);
            ws.Range("B11:B20").CreateDataValidation().Time.NotEqualTo(c1);
            ws.Range("C11:C20").CreateDataValidation().Time.GreaterThan(c1);
            ws.Range("D11:D20").CreateDataValidation().Time.LessThan(c1);
            ws.Range("E11:E20").CreateDataValidation().Time.EqualOrGreaterThan(c1);
            ws.Range("F11:F20").CreateDataValidation().Time.EqualOrLessThan(c1);
            ws.Range("G11:G20").CreateDataValidation().Time.Between(c1, c2);
            ws.Range("H11:H20").CreateDataValidation().Time.NotBetween(c1, c2);

            wb.SaveAs(filePath);
        }
    }
}
