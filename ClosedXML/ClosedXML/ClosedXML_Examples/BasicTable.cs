using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.Excel.Style;

namespace ClosedXML_Examples
{
    public class BasicTable
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Contacts");

            //First Names
            ws.Cell("A1").Value = "FName";
            ws.Cell("A2").Value = "John";
            ws.Cell("A3").Value = "Hank";
            ws.Cell("A4").Value = "Dagny";
            //Last Names
            ws.Cell("B1").Value = "LName";
            ws.Cell("B2").Value = "Galt";
            ws.Cell("B3").Value = "Rearden";
            ws.Cell("B4").Value = "Taggart";
            //Is an outcast?
            ws.Cell("C1").Value = "Outcast";
            ws.Cell("C2").Value = true.ToString();
            ws.Cell("C3").Value = false.ToString();
            ws.Cell("C4").Value = false.ToString();
            //Date of Birth
            ws.Cell("D1").Value = "DOB";
            ws.Cell("D2").Value = new DateTime(1919, 1, 21).ToString();
            ws.Cell("D3").Value = new DateTime(1907, 3, 4).ToString();
            ws.Cell("D4").Value = new DateTime(1921, 12, 15).ToString();
            //Income
            ws.Cell("E1").Value = "Income";
            ws.Cell("E2").Value = "2000";
            ws.Cell("E3").Value = "40000";
            ws.Cell("E4").Value = "10000";

            //var rngDates = ws.Range("D2:D4");
            //var rngNumbers = ws.Range("E2:E4");

            //rngDates.Style.NumberFormat.Format = "mm-dd-yy";
            //rngNumbers.Style.NumberFormat.Format = "$ #,##0";

            //var rngHeaders = ws.Range("A1:E1");
            //rngHeaders.Style.Font.Bold = true;
            //rngHeaders.Style.Fill.BackgroundColor = "6BE8FF";

            //var rngTable = ws.Range("A1:E4");
            //rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

            workbook.SaveAs(filePath);
        }
    }
}
