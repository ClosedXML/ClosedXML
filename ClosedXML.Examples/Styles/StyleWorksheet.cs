using ClosedXML.Excel;

namespace ClosedXML.Examples.Styles
{
    public class StyleWorksheet : IXLExample
    {
        public void Create(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Style Worksheet");

            ws.Style.Font.Bold = true;
            ws.Style.Font.FontColor = XLColor.Red;
            ws.Style.Fill.BackgroundColor = XLColor.Cyan;

            // The following cells will be bold and red
            // because we've specified those attributes to the entire worksheet
            ws.Cell(1, 1).Value = "Test";
            ws.Cell(1, 2).Value = "Case";

            // Here we'll change the style of a single cell
            ws.Cell(2, 1).Value = "Default";
            ws.Cell(2, 1).Style = XLWorkbook.DefaultStyle;

            // Let's play with some rows
            ws.Row(4).Style = XLWorkbook.DefaultStyle;
            ws.Row(4).Height = 20;
            ws.Rows(5, 6).Style = XLWorkbook.DefaultStyle;
            ws.Rows(5, 6).Height = 20;

            // Let's play with some columns
            ws.Column(4).Style = XLWorkbook.DefaultStyle;
            ws.Column(4).Width = 5;
            ws.Columns(5, 6).Style = XLWorkbook.DefaultStyle;
            ws.Columns(5, 6).Width = 5;

            workbook.SaveAs(filePath);
        }
    }
}