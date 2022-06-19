using ClosedXML.Excel;

namespace ClosedXML.Examples.Styles
{
    public class StyleRowsColumns : IXLExample
    {
        public void Create(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Style Rows and Columns");

            //Set the entire worksheet's cells to be bold and with a light cyan background
            ws.Style.Font.Bold = true;
            ws.Style.Fill.BackgroundColor = XLColor.LightCyan;

            // Set the width of all columns in the worksheet
            ws.Columns().Width = 5;

            // Set the height of all rows in the worksheet
            ws.Rows().Height = 20;

            // Let's play with the rows and columns
            ws.Rows(2, 3).Style.Fill.BackgroundColor = XLColor.Blue;
            ws.Columns(3, 4).Style.Fill.BackgroundColor = XLColor.Orange;
            ws.Rows(5, 5).Style.Fill.BackgroundColor = XLColor.Pink;
            ws.Row(6).Style.Fill.BackgroundColor = XLColor.Brown;
            ws.Column("E").Style.Fill.BackgroundColor = XLColor.Gray;

            workbook.SaveAs(filePath);
        }
    }
}