using ClosedXML.Excel;
using System.IO;

namespace ClosedXML.Examples
{
    public class ShiftingRanges : IXLExample
    {
        public void Create(string filePath)
        {
            var tempFile = ExampleHelper.GetTempFilePath(filePath);
            try
            {
                new BasicTable().Create(tempFile);
                using var workbook = new XLWorkbook(tempFile);
                var ws = workbook.Worksheet(1);

                // Get a range object
                var rngHeaders = ws.Range("B3:F3");

                // Insert some rows/columns before the range
                ws.Row(1).InsertRowsAbove(2);
                ws.Column(1).InsertColumnsBefore(2);

                // Change the background color of the headers
                rngHeaders.Style.Fill.BackgroundColor = XLColor.LightSalmon;

                ws.Columns().AdjustToContents();

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