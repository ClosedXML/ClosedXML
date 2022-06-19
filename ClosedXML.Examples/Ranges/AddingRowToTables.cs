using ClosedXML.Excel;
using System.IO;
using System.Linq;

namespace ClosedXML.Examples.Ranges
{
    public class AddingRowToTables : IXLExample
    {
        #region Methods

        // Public
        public void Create(string filePath)
        {
            var tempFile = ExampleHelper.GetTempFilePath(filePath);
            try
            {
                new BasicTable().Create(tempFile);
                using var wb = new XLWorkbook(tempFile);
                var ws = wb.Worksheets.First();

                var firstCell = ws.FirstCellUsed();
                var lastCell = ws.LastCellUsed();
                var range = ws.Range(firstCell.Address, lastCell.Address);
                range.FirstRow().Delete(); // Deleting the "Contacts" header (we don't need it for our purposes)

                // We want to use a theme for table, not the hard coded format of the BasicTable
                range.Clear(XLClearOptions.AllFormats);
                // Put back the date and number formats
                range.Column(4).Style.NumberFormat.NumberFormatId = 15;
                range.Column(5).Style.NumberFormat.Format = "$ #,##0";

                var table = range.CreateTable(); // You can also use range.AsTable() if you want to

                ws.Cell("Q6000").Value = "dummy value";

                var row = table.DataRange.InsertRowsBelow(1).First();

                wb.SaveAs(filePath);
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }

        // Private

        // Override

        #endregion Methods
    }
}