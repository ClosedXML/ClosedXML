using ClosedXML.Excel;

namespace ClosedXML.Examples.Styles
{
    public class DefaultStyles : IXLExample
    {
        public void Create(string filePath)
        {
            // Create our workbook
            using var workbook = new XLWorkbook();

            // This worksheet will have the default style, row height, column width, and page setup
            var ws1 = workbook.Worksheets.Add("Default Style");

            // Change the default row height for all new worksheets in this workbook
            workbook.RowHeight = 30;

            var ws2 = workbook.Worksheets.Add("Tall Rows");

            // Create a worksheet and change the default row height
            var ws3 = workbook.Worksheets.Add("Short Rows");
            ws3.RowHeight = 7.5;

            workbook.SaveAs(filePath);
        }
    }
}