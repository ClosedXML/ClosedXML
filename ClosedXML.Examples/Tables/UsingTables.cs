using ClosedXML.Excel;
using System.IO;

namespace ClosedXML.Examples.Tables
{
    public class UsingTables : IXLExample
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
                var ws = wb.Worksheet(1);
                ws.Name = "Contacts Table";
                var firstCell = ws.FirstCellUsed();
                var lastCell = ws.LastCellUsed();
                var range = ws.Range(firstCell.Address, lastCell.CellRight().Address);
                range.FirstRow().Delete(); // Deleting the "Contacts" header (we don't need it for our purposes)

                // We want to use a theme for table, not the hard coded format of the BasicTable
                range.Clear(XLClearOptions.AllFormats);
                // Put back the date and number formats
                range.Column(4).Style.NumberFormat.NumberFormatId = 15;
                range.Column(5).Style.NumberFormat.Format = "$ #,##0";

                // Add a field
                range.Column(6).FirstCell().SetValue("Age");
                var c = range.Column(6).FirstCell().CellBelow();
                c.Style.NumberFormat.SetFormat("0.00");
                c.FormulaA1 = "=(DATE(2017, 10, 3) - E3) / 365";

                c.CopyTo(c.CellBelow())
                    .CopyTo(c.CellBelow().CellBelow());

                var table = range.CreateTable(); // You can also use range.AsTable() if you want to
                                                 // manipulate the range as a table but don't want
                                                 // to create the table in the worksheet.

                // Let's activate the Totals row and add the sum of Income
                table.ShowTotalsRow = true;
                table.Field("Income").TotalsRowFunction = XLTotalsRowFunction.Sum;
                // Just for fun let's add the text "Sum Of Income" to the totals row
                table.Field(0).TotalsRowLabel = "Sum Of Income";

                table.Field("Age").TotalsRowFunction = XLTotalsRowFunction.Average;

                // Copy all the headers
                var columnWithHeaders = lastCell.Address.ColumnNumber + 3;
                var currentRow = table.RangeAddress.FirstAddress.RowNumber;
                ws.Cell(currentRow, columnWithHeaders).Value = "Table Headers";
                foreach (var cell in table.HeadersRow().Cells())
                {
                    currentRow++;
                    ws.Cell(currentRow, columnWithHeaders).Value = cell.Value;
                }

                // Format the headers as a table with a different style and no autofilters
                var htFirstCell = ws.Cell(table.RangeAddress.FirstAddress.RowNumber, columnWithHeaders);
                var htLastCell = ws.Cell(currentRow, columnWithHeaders);
                var headersTable = ws.Range(htFirstCell, htLastCell).CreateTable("Headers");
                headersTable.Theme = XLTableTheme.TableStyleLight10;
                headersTable.ShowAutoFilter = false;

                // Add a custom formula to the headersTable
                headersTable.ShowTotalsRow = true;
                headersTable.Field(0).TotalsRowFormulaA1 = "CONCATENATE(\"Count: \", CountA(Headers[Table Headers]))";

                // Copy the names
                var columnWithNames = columnWithHeaders + 2;
                currentRow = table.RangeAddress.FirstAddress.RowNumber; // reset the currentRow
                ws.Cell(currentRow, columnWithNames).Value = "Names";
                foreach (var row in table.DataRange.Rows())
                {
                    currentRow++;
                    var fName = row.Field("FName").GetString(); // Notice how we're calling the cell by field name
                    var lName = row.Field("LName").GetString(); // Notice how we're calling the cell by field name
                    var name = string.Format("{0} {1}", fName, lName);
                    ws.Cell(currentRow, columnWithNames).Value = name;
                }

                // Format the names as a table with a different style and no autofilters
                var ntFirstCell = ws.Cell(table.RangeAddress.FirstAddress.RowNumber, columnWithNames);
                var ntLastCell = ws.Cell(currentRow, columnWithNames);
                var namesTable = ws.Range(ntFirstCell, ntLastCell).CreateTable();
                namesTable.Theme = XLTableTheme.TableStyleLight12;
                namesTable.ShowAutoFilter = false;

                ws.Columns().AdjustToContents();
                ws.Columns("A,H,J").Width = 3;

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
