using System.IO;
using ClosedXML.Excel;


namespace ClosedXML.Examples
{
    public class ChangingBasicTable : IXLExample
    {
        public void Create(string filePath)
        {
            string tempFile = ExampleHelper.GetTempFilePath(filePath);
            try
            {
                new BasicTable().Create(tempFile);
                var workbook = new XLWorkbook(tempFile);
                var ws = workbook.Worksheet(1);

                // Change the background color of the headers
                var rngHeaders = ws.Range("B3:F3");
                rngHeaders.Style.Fill.BackgroundColor = XLColor.LightSalmon;

                // Change the date formats
                var rngDates = ws.Range("E4:E6");
                rngDates.Style.DateFormat.Format = "MM/dd/yyyy";

                // Change the income values to text
                var rngNumbers = ws.Range("F4:F6");
                foreach (var cell in rngNumbers.Cells())
                {
                    string formattedString = cell.GetFormattedString();
                    cell.DataType = XLDataType.Text;
                    cell.Value = formattedString + " Dollars";
                }

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