using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Drawing;


namespace ClosedXML_Examples
{
    public class ChangingBasicTable
    {
        public void Create()
        {
            var workbook = new XLWorkbook(@"C:\Excel Files\Created\BasicTable.xlsx");
            var ws = workbook.Worksheets.Worksheet(0);

            // Change the background color of the headers
            var rngHeaders = ws.Range("B3:F3");
            rngHeaders.Style.Fill.BackgroundColor = Color.LightSalmon;

            // Change the date formats
            var rngDates = ws.Range("E4:E6");
            rngDates.Style.DateFormat.Format = "MM/dd/yyyy";

            // Change the income values to text
            var rngNumbers = ws.Range("F4:F6");
            foreach (var cell in rngNumbers.Cells())
            {
                cell.DataType = XLCellValues.Text;
                cell.Value += " Dollars";
            }

            ws.Columns().AdjustToContents();

            workbook.SaveAs(@"C:\Excel Files\Created\BasicTable_Modified.xlsx");
        }
    }
}