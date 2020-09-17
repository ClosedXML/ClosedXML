using ClosedXML.Excel;
using System;

namespace ClosedXML_Examples.Misc
{
    public class DataValidationTime : IXLExample
    {
        public void Create(string filePath)
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var time1 = TimeSpan.FromHours(1);
            var time2 = TimeSpan.FromMinutes(150);
            var c1 = ws.Cell("A1");
            var c2 = ws.Cell("B1");
            c1.Value = time1;
            c2.Value = time2;

            ws.Range("A2:A10").SetDataValidation().Time.EqualTo(time1);
            ws.Range("B2:B10").SetDataValidation().Time.NotEqualTo(time1);
            ws.Range("C2:C10").SetDataValidation().Time.GreaterThan(time1);
            ws.Range("D2:D10").SetDataValidation().Time.LessThan(time1);
            ws.Range("E2:E10").SetDataValidation().Time.EqualOrGreaterThan(time1);
            ws.Range("F2:F10").SetDataValidation().Time.EqualOrLessThan(time1);
            ws.Range("G2:G10").SetDataValidation().Time.Between(time1, time2);
            ws.Range("H2:H10").SetDataValidation().Time.NotBetween(time1, time2);

            ws.Range("A11:A20").SetDataValidation().Time.EqualTo(c1);
            ws.Range("B11:B20").SetDataValidation().Time.NotEqualTo(c1);
            ws.Range("C11:C20").SetDataValidation().Time.GreaterThan(c1);
            ws.Range("D11:D20").SetDataValidation().Time.LessThan(c1);
            ws.Range("E11:E20").SetDataValidation().Time.EqualOrGreaterThan(c1);
            ws.Range("F11:F20").SetDataValidation().Time.EqualOrLessThan(c1);
            ws.Range("G11:G20").SetDataValidation().Time.Between(c1, c2);
            ws.Range("H11:H20").SetDataValidation().Time.NotBetween(c1, c2);

            wb.SaveAs(filePath);
        }
    }
}