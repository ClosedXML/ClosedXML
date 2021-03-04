using System.IO;
using ClosedXML.Excel;

namespace ClosedXML.Examples
{
    public class CopyingRanges : IXLExample
    {
        public void Create(string filePath)
        {
            var tempFile = ExampleHelper.GetTempFilePath(filePath);
            try
            {
                new BasicTable().Create(tempFile);
                var workbook = new XLWorkbook(tempFile);
                var ws = workbook.Worksheet(1);

                // Define a range with the data
                var firstTableCell = ws.FirstCellUsed();
                var lastTableCell = ws.LastCellUsed();
                var rngData = ws.Range(firstTableCell.Address, lastTableCell.Address);

                // Copy the table to another worksheet
                var wsCopy = workbook.Worksheets.Add("Contacts Copy");
                wsCopy.Cell(1, 1).Value = rngData;

                workbook.SaveAs(filePath);
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }
    }
}