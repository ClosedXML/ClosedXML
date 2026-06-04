using NUnit.Framework;

namespace ClosedXML.Tests.Excel;

[TestFixture]
internal class TableStyleTests
{
    [Test]
    public void Load_and_save_field_differential_styles()
    {
        // Test file contains different dxf for header, data and totals of each column. The table
        // doesn't have a header, because dxf for header can be specified only when header is not
        // shown. Toggle Header Row to see the format of the header in Excel. The table should
        // have font size that increases in each table area, from 10 to 15 points.
        TestHelper.LoadSaveAndCompare(
            @"Other\Tables\TableColumnStyles-input.xlsx",
            @"Other\Tables\TableColumnStyles-output.xlsx");
    }

    [Test]
    public void Load_and_save_table_with_table_style()
    {
        // Test file contains a table style with a different dxf for every region other than
        // WholeTable. WholeTable region is omitted to test that omitting works and doesn't
        // write some dxf even for a region without dxf. The test file contains three tables
        // to demonstrate the application of the style. That is needed, because some
        // combinations are meaningless in one table (e.g. styling of row and column stripes
        // at once).
        TestHelper.LoadModifyAndCompare(
            @"Other\Tables\TableStyle-input.xlsx",
            wb =>
            {
                var ws = wb.Worksheet(1);
                ws.Cell("A10").Value = "Edited";
            },
            @"Other\Tables\TableStyle-output.xlsx");
    }
}
