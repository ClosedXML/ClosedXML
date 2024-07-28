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
}
