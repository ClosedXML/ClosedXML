using ClosedXML.Excel;
using System;

namespace ClosedXML.Examples
{
    public class BasicTable : IXLExample
    {
        public void Create(string filePath)
        {
            // Creating a new workbook
            var wb = new XLWorkbook();

            //Adding a worksheet
            var ws = wb.Worksheets.Add("Contacts");

            //Adding text
            //Title
            ws.Cell("B2").Value = "Contacts";
            //First Names
            ws.Cell("B3").Value = "FName";
            ws.Cell("B4").Value = "John";
            ws.Cell("B5").Value = "Hank";
            ws.Cell("B6").Value = "Dagny";
            //Last Names
            ws.Cell("C3").Value = "LName";
            ws.Cell("C4").Value = "Galt";
            ws.Cell("C5").Value = "Rearden";
            ws.Cell("C6").Value = "Taggart";

            //Adding more data types
            //Is an outcast?
            ws.Cell("D3").Value = "Outcast";
            ws.Cell("D4").Value = true;
            ws.Cell("D5").Value = false;
            ws.Cell("D6").Value = false;
            //Date of Birth
            ws.Cell("E3").Value = "DOB";
            ws.Cell("E4").Value = new DateTime(1919, 1, 21);
            ws.Cell("E5").Value = new DateTime(1907, 3, 4);
            ws.Cell("E6").Value = new DateTime(1921, 12, 15);
            //Income
            ws.Cell("F3").Value = "Income";
            ws.Cell("F4").Value = 2000;
            ws.Cell("F5").Value = 40000;
            ws.Cell("F6").Value = 10000;

            //Defining ranges
            //From worksheet
            var rngTable = ws.Range("B2:F6");
            //From another range
            var rngDates = rngTable.Range("E4:E6");
            var rngNumbers = rngTable.Range("F4:F6");

            //Formatting dates and numbers
            //Using a OpenXML's predefined formats
            rngDates.Style.NumberFormat.NumberFormatId = 15;
            //Using a custom format
            rngNumbers.Style.NumberFormat.Format = "$ #,##0";

            //Formatting headers
            var rngHeaders = rngTable.Range("B3:F3");
            rngHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rngHeaders.Style.Font.Bold = true;
            rngHeaders.Style.Fill.BackgroundColor = XLColor.Aqua;

            //Adding grid lines
            rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

            //Format title cell
            rngTable.Cell(1, 1).Style.Font.Bold = true;
            rngTable.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.CornflowerBlue;
            rngTable.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            //Merge title cells
            rngTable.Row(1).Merge(); // We could've also used: rngTable.Range("A1:E1").Merge()

            //Add thick borders
            rngTable.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

            // You can also specify the border for each side with:
            // rngTable.FirstColumn().Style.Border.LeftBorder = XLBorderStyleValues.Thick;
            // rngTable.LastColumn().Style.Border.RightBorder = XLBorderStyleValues.Thick;
            // rngTable.FirstRow().Style.Border.TopBorder = XLBorderStyleValues.Thick;
            // rngTable.LastRow().Style.Border.BottomBorder = XLBorderStyleValues.Thick;

            // Adjust column widths to their content
            ws.Columns(2, 6).AdjustToContents();

            //Saving the workbook
            wb.SaveAs(filePath);
        }
    }
}
