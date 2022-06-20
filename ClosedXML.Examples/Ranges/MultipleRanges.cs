using ClosedXML.Excel;

namespace ClosedXML.Examples
{
    public class MultipleRanges : IXLExample
    {
        public void Create(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Multiple Ranges");

            // using multiple string range definitions
            ws.Ranges("A1:B2,C3:D4,E5:F6").Style.Fill.BackgroundColor = XLColor.Red;

            // using a single string separated by commas
            ws.Ranges("A5:B6,E1:F2").Style.Fill.BackgroundColor = XLColor.Orange;

            workbook.SaveAs(filePath);
        }
    }
}