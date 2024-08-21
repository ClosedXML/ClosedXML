using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.PivotTables.Style;

internal class XLPivotFieldStyleFormatsTests
{
    [Test]
    public void Modify_pivot_field_label_style()
    {
        TestHelper.CreateAndCompare(wb =>
        {
            var dataSheet = wb.AddWorksheet();
            var dataRange = dataSheet.Cell("A1").InsertData(new object[]
            {
                ("Name", "Month", "Price"),
                ("Cake", "Jan", 9),
                ("Pie", "Jan", 7),
                ("Cake", "Feb", 3),
            });

            var ptSheet = wb.AddWorksheet().SetTabActive();
            ptSheet.Column("A").Width = 15;
            var pt = dataRange.CreatePivotTable(ptSheet.Cell("A1"), "pivot table");
            pt.RowLabels.Add("Name");
            var monthField = pt.RowLabels.Add("Month");
            pt.Values.Add("Price");

            // Modify style in two steps to ensure second modification doesn't override the first one
            monthField.StyleFormats.Label.Style
                .Fill.SetBackgroundColor(XLColor.LightGray)
                .Font.SetStrikethrough();
            monthField.StyleFormats.Label.Style.Font.SetBold();
        }, @"Other\PivotTable\Style\Modify_pivot_field_label_style.xlsx");
    }

    [TestCase(XLPivotLayout.Compact, "Set_pivot_field_header_style-compact.xlsx")]
    [TestCase(XLPivotLayout.Tabular, "Set_pivot_field_header_style-tabular.xlsx")]
    public void Set_pivot_field_header_style(XLPivotLayout layout, string testFile)
    {
        // Header in compact is only one cell, whereas tabular has individual header for each field
        // on axis. Tested axis contains two fields to check that even when there is only one header,
        // it is used for all fields (i.e. the single header cell is colored, not a cell next to it).
        TestHelper.CreateAndCompare(wb =>
        {
            var dataSheet = wb.AddWorksheet();
            var dataRange = dataSheet.Cell("A1").InsertData(new object[]
            {
                ("Name", "Flavor", "Month", "Price"),
                ("Cake", "Vanilla", "Jan", 9),
                ("Pie", "Peach", "Jan", 7),
                ("Cake", "Lemon", "Feb", 3),
            });

            var ptSheet = wb.AddWorksheet().SetTabActive();
            var pt = dataRange.CreatePivotTable(ptSheet.Cell("A1"), "pivot table");
            pt.Layout = layout;
            pt.Values.Add("Price");
            pt.RowLabels.Add("Month");
            pt.ColumnLabels.Add("Name");
            var styledHeaderField = pt.ColumnLabels.Add("Flavor");

            // Set two style in two steps to check that second one doesn't overwrite first one.
            styledHeaderField.StyleFormats.Header.Style.Fill.SetBackgroundColor(XLColor.Green);
            styledHeaderField.StyleFormats.Header.Style.Font.SetFontColor(XLColor.Red);
        }, $@"Other\PivotTable\Style\{testFile}");
    }

    [Test]
    public void Set_pivot_field_subtotals_style()
    {
        // In the test we set two subtotals, one for name and other for month. The month one
        // is there to check the subtotal of a last field on an multi-field axis is displayed
        // correctly (it needs Outline:0 attribute to be displayed correctly in Excel).
        TestHelper.CreateAndCompare(wb =>
        {
            var dataSheet = wb.AddWorksheet();
            var dataRange = dataSheet.Cell("A1").InsertData(new object[]
            {
                ("Name", "Month", "Price"),
                ("Cake", "Jan", 9),
                ("Pie", "Jan", 7),
                ("Cake", "Feb", 3),
            });

            var ptSheet = wb.AddWorksheet().SetTabActive();
            var pt = dataRange.CreatePivotTable(ptSheet.Cell("A1"), "pivot table");
            pt.Values.Add("Price");
            var nameField = pt.RowLabels.Add("Name")
                .AddSubtotal(XLSubtotalFunction.Sum)
                .SetSubtotalsAtTop(false);
            var monthField = pt.RowLabels.Add("Month")
                .AddSubtotal(XLSubtotalFunction.Maximum)
                .AddSubtotal(XLSubtotalFunction.Minimum)
                .SetSubtotalsAtTop(false);

            // Set two style in two steps to check that second one doesn't overwrite first one
            // and also that second access modifies the same pivot area.
            nameField.StyleFormats.Subtotal.Style.Font.SetFontColor(XLColor.GoldenBrown);
            nameField.StyleFormats.Subtotal.Style.Fill.SetBackgroundColor(XLColor.Yellow);

            monthField.StyleFormats.Subtotal.Style.Font.SetFontColor(XLColor.Green).Font.SetBold();
            monthField.StyleFormats.Subtotal.Style.Fill.SetBackgroundColor(XLColor.Red);
        }, @"Other\PivotTable\Style\Set_pivot_field_subtotals_style.xlsx");
    }
}
