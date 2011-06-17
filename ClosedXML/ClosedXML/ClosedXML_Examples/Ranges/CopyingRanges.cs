using ClosedXML.Excel;

namespace ClosedXML_Examples
{
    public class CopyingRanges
    {
        public void Create()
        {
            var workbook = new XLWorkbook(@"C:\Excel Files\Created\BasicTable.xlsx");
            var ws = workbook.Worksheet(1);

            // Define a range with the data
            var firstTableCell = ws.FirstCellUsed();
            var lastTableCell = ws.LastCellUsed();
            var rngData = ws.Range(firstTableCell.Address, lastTableCell.Address);

            // Copy the table to another worksheet
            var wsCopy = workbook.Worksheets.Add("Contacts Copy");
            wsCopy.Cell(1,1).Value = rngData;

            workbook.SaveAs(@"C:\Excel Files\Created\CopyingRanges.xlsx");
        }
    }
}