using System.Linq;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.IO;

[TestFixture]
internal class WorksheetReaderTests
{
    [Test]
    public void Keep_track_of_row_number_even_if_omitted_in_row_element()
    {
        // The test file relies on a tracked row number to position cells and rows. When row:r is
        // not specified, it should be the next row after the last read row.
        TestHelper.LoadAndAssert((_, ws) =>
        {
            // Row 3 contains an explicit value that is same as the tracked value.
            // There is a gap at rows 5-6 to ensure tracked value detects the gap.

            // Cells that depend on tracked row number have their addresses determined correctly
            foreach (var row in new[] { 1, 2, 3, 4, 7, 8 })
                Assert.AreEqual(row, ws.Cell(row, 1).Value);

            // Check that rows numbers determined correctly. Row 9 doesn't have any cells
            foreach (var row in new[] { 1, 2, 3, 4, 7, 8, 9 })
                Assert.AreEqual(40 + row, ws.Rows().Single(x => x.RowNumber() == row).Height);
        }, "Other.IO.Worksheet.OmittedRowNumber.xlsx");
    }

    [Test]
    public void Keep_track_of_cell_address_even_if_omitted_in_c_element()
    {
        // The test file relies on a tracked column number to position cells. If a cell doesn't
        // specify c:r, it should be a cell to the right of the last read cell.
        TestHelper.LoadAndAssert((_, ws) =>
        {
            // The C1 is redundantly specified. There is a gap at columns E-F to ensure
            // tracking restarts after last know cell ref.
            foreach (var column in new[] { "A", "B", "C", "D", "G", "H" })
                Assert.AreEqual(column, ws.Cell($"{column}1").Value);

            // When reader encounters a new row, the tracked tracked colum is reset back to 'A'
            foreach (var column in new[] { "A", "B" })
                Assert.AreEqual(column, ws.Cell($"{column}2").Value);
        }, "Other.IO.Worksheet.OmittedCellReference.xlsx");
    }
}
