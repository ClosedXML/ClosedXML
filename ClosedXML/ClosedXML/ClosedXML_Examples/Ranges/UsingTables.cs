using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;

namespace ClosedXML_Examples.Ranges
{
    public class UsingTables
    {
        #region Methods

        // Public
        public void Create(String filePath)
        {
            var wb = new XLWorkbook(@"C:\Excel Files\Created\BasicTable.xlsx");
            var ws = wb.Worksheets.Worksheet(0);
            var firstCell = ws.FirstCellUsed();
            var lastCell = ws.LastCellUsed();
            var range = ws.Range(firstCell.Address, lastCell.Address);
            range.Row(1).Delete(); // Deleting the "Contacts" header (we don't need it for our purposes)
            range.ClearStyles(); // We want to use a theme for table, not the hard coded format of the BasicTable

            var table = range.CreateTable();    // You can also use range.AsTable() if you want to
                                                // manipulate the range as a table but don't want 
                                                // to create the table in the worksheet.

            // Let's activate the Totals row and add the sum of Income
            table.ShowTotalsRow = true;
            table.Field("Income").TotalsRowFunction = XLTotalsRowFunction.Sum;
            // Just for fun let's add the text "Sum Of Income" to the totals row
            table.Field(0).TotalsRowLabel = "Sum Of Income";

            // Copy all the headers
            Int32 columnWithHeaders = lastCell.Address.ColumnNumber + 2;
            Int32 currentRow = table.RangeAddress.FirstAddress.RowNumber;
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
            Int32 columnWithNames = columnWithHeaders + 2;
            currentRow = table.RangeAddress.FirstAddress.RowNumber; // reset the currentRow
            ws.Cell(currentRow, columnWithNames).Value = "Names";
            foreach (var row in table.Rows())
            {
                currentRow++;
                var fName = row.Field("FName").GetString(); // Notice how we're calling the cell by field name
                var lName = row.Field("LName").GetString(); // Notice how we're calling the cell by field name
                var name = String.Format("{0} {1}", fName, lName);
                ws.Cell(currentRow, columnWithNames).Value = name;
            }

            // Format the names as a table with a different style and no autofilters
            var ntFirstCell = ws.Cell(table.RangeAddress.FirstAddress.RowNumber, columnWithNames);
            var ntLastCell = ws.Cell(currentRow, columnWithNames);
            var namesTable = ws.Range(ntFirstCell, ntLastCell).CreateTable();
            namesTable.Theme = XLTableTheme.TableStyleLight12;
            namesTable.ShowAutoFilter = false;

            ws.Columns().AdjustToContents();
            ws.Columns("A,G,I").Width = 3;

            wb.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
