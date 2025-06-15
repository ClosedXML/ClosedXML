using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine.Visitors;
using ClosedXML.Extensions;
using ClosedXML.Parser;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine;

[TestFixture]
[TestOf(typeof(ReferenceShiftOnInsertRefModVisitor))]
[TestOf(typeof(ReferenceShiftOnDeleteRefModVisitor))]
public class ReferenceShiftRefModTests
{
    [TestCase("C2", "Other!C2", "C2")]
    [TestCase("$C2", "Sheet!C2", "$C3")]
    [TestCase("C2", "Sheet!B2", "C2")]
    [TestCase("C2", "Sheet!D2", "C2")]
    [TestCase("C2", "Sheet!C3", "C2")]
    [TestCase("C2", "Sheet!C1", "C3")]
    [TestCase("C2:D4", "Sheet!C1:D3", "C5:D7")]
    [TestCase("C2:D4", "Sheet!C2:D2", "C3:D5")]
    [TestCase("C2:D4", "Sheet!C3:D4", "C2:D6")]
    [TestCase("C2:D4", "Sheet!D1", "C2:D4")] // Would cause split
    [TestCase("A1048576", "Sheet!A1048576", "#REF!")]

    // Reference with sheet is modified to a reference with a sheet
    [TestCase("Sheet1!A1", "Sheet1!A1", "Sheet1!A2")]
    [TestCase("'Joe''s Bakery'!C2", "'Joe''s Bakery'!C2", "'Joe''s Bakery'!C3")]
    public void Insert_area_and_shift_down_reference(string formula, string insertedArea, string expected)
    {
        // TODO: Once incorporated into area insertion, replace with a public API test case through SUM(reference) in a cell.
        Assert.True(ReferenceParser.TryParseSheetA1(insertedArea, out var insertedSheet, out var insertedReference));
        var inserted = new XLBookArea(insertedSheet, insertedReference.ToSheetRangeA1());

        var result = FormulaConverter.ModifyA1(formula, "Sheet", 1, 1, new ReferenceShiftOnInsertRefModVisitor(inserted, true));

        Assert.AreEqual(expected, result);
    }

    [TestCase("C2", "Sheet!B2", "D2")]
    [TestCase("$C$2", "Sheet!B2", "$D$2")]
    [TestCase("C2", "Sheet!C1", "C2")]
    [TestCase("C2", "Sheet!D1", "C2")]
    [TestCase("C2", "Sheet!E1", "C2")]
    [TestCase("C2", "Sheet!E2", "C2")]
    [TestCase("C2", "Sheet!E3", "C2")]
    [TestCase("C2", "Sheet!D3", "C2")]
    [TestCase("C2", "Sheet!B3", "C2")]
    [TestCase("C2:E5", "Sheet!E2:G5", "C2:H5")] // Insert into middle of area
    [TestCase("XFD1", "Sheet!XFD1", "#REF!")] // Pushed out of a sheet
    [TestCase("XFB1:XFC1", "Sheet!XFC1:XFD1", "XFB1:XFD1")] // Grew out of a sheet
    [TestCase("4:$5", "Sheet!B1:E10", "4:$5")]
    [TestCase("$D:E", "Sheet!B1:E5", "$D:E")]

    // Would cause split
    [TestCase("C2:E5", "Sheet!B2:B3", "C2:E5")]
    [TestCase("C2:E5", "Sheet!C2", "C2:E5")]
    [TestCase("C2:E5", "Sheet!D2:D3", "C2:E5")]

    // Reference with sheet is modified to a reference with a sheet
    [TestCase("Sheet1!C2:F2", "Sheet1!D2", "Sheet1!C2:G2")]
    [TestCase("'Joe''s Bakery'!C5:D5", "'Joe''s Bakery'!B5:C5", "'Joe''s Bakery'!E5:F5")]
    public void Insert_area_and_shift_right_reference(string formula, string insertedArea, string expected)
    {
        // TODO: Once incorporated into area insertion, replace with a public API test case through SUM(reference) in a cell.
        Assert.True(ReferenceParser.TryParseSheetA1(insertedArea, out var insertedSheet, out var insertedReference));
        var inserted = new XLBookArea(insertedSheet, insertedReference.ToSheetRangeA1());

        var result = FormulaConverter.ModifyA1(formula, "Sheet", 1, 1, new ReferenceShiftOnInsertRefModVisitor(inserted, false));

        Assert.AreEqual(expected, result);
    }

    /// <summary>
    /// This tests how are references changed inside a formula when an area is deleted and shifted up.
    /// They were derived by using SUM(reference) before and after deletion.
    /// </summary>
    [TestCase("C2", "Other!C2", "C2")] // Delete cell with same coordinate on different sheet.
    [TestCase("C2", "Sheet!C2", "#REF!")] // Fully delete a cell.
    [TestCase("$C$2", "Sheet!C2", "#REF!")] // Only posıtıon matter, reference style doesn't.

    // Column span tests
    [TestCase("A:B", "Sheet!10:20", "A:B")] // Column span is basically never changed.
    [TestCase("A:B", "Sheet!A1:B5", "A:B")] // Column span is basically never changed.

    // Deleted area is fully to the left of a reference.
    [TestCase("C1:D3", "Sheet!A1:B3", "C1:D3")]

    // Deleted area is fully to the right of a reference.
    [TestCase("A1:B3", "Sheet!C1:D3", "A1:B3")]

    // Deleted area is inside and only partially covers a reference.
    [TestCase("A1:D3", "Sheet!B1:C3", "A1:D3")]

    // Deleted area covers full width of reference
    [TestCase("A1:D3", "Sheet!A5:D6", "A1:D3")] // Deleted area is completely below the reference.
    [TestCase("A5:A10", "Sheet!A10:D11", "A5:A9")]
    [TestCase("A5:A10", "Sheet!A10", "A5:A9")]
    [TestCase("A5:A10", "Sheet!A7:A10", "A5:A6")]
    [TestCase("A5:A10", "Sheet!A5:A10", "#REF!")]
    [TestCase("A5:A10", "Sheet!A5:A6", "A5:A8")]
    [TestCase("A5:A10", "Sheet!A4:A6", "A4:A7")]
    [TestCase("A5:A10", "Sheet!A2:A4", "A2:A7")]

    // Deleted area covers a left slice of a reference
    [TestCase("C5:E10", "Sheet!C5:D10", "E5:E10")]
    [TestCase("C5:E10", "Sheet!A5:D10", "E5:E10")]
    [TestCase("C5:E10", "Sheet!A2:D15", "E5:E10")]
    [TestCase("C5:E10", "Sheet!A6:D10", "C5:E10")]
    [TestCase("C5:E10", "Sheet!A5:D9", "C5:E10")]
    [TestCase("C5:E10", "Sheet!C1:D4", "C5:E10")]
    [TestCase("C5:E10", "Sheet!C11:D14", "C5:E10")]

    // Deleted area covers a right slice of a reference
    [TestCase("C5:E10", "Sheet!D5:E10", "C5:C10")]
    [TestCase("C5:E10", "Sheet!D5:F10", "C5:C10")]
    [TestCase("C5:E10", "Sheet!D2:E15", "C5:C10")]
    [TestCase("C5:E10", "Sheet!D6:F10", "C5:E10")]
    [TestCase("C5:E10", "Sheet!D5:F9", "C5:E10")]
    [TestCase("C5:E10", "Sheet!D1:E4", "C5:E10")]
    [TestCase("C5:E10", "Sheet!D11:E14", "C5:E10")]
    public void Delete_area_and_shift_up_reference(string formula, string deletedArea, string expected)
    {
        // TODO: Once incorporated into cell deletion, replace with a public API test case through SUM(reference) in a cell.
        Assert.True(ReferenceParser.TryParseSheetA1(deletedArea, out var deletedSheet, out var deletedReference));
        var deleted = new XLBookArea(deletedSheet, deletedReference.ToSheetRangeA1());

        var result = FormulaConverter.ModifyA1(formula, "Sheet", 1, 1, new ReferenceShiftOnDeleteRefModVisitor(deleted, XLShiftDeletedCells.ShiftCellsUp));

        Assert.AreEqual(expected, result);
    }

    /// <summary>
    /// This tests how are references changed inside a formula when an area is deleted and shifted left.
    /// They were derived by using SUM(reference) before and after deletion.
    /// </summary>
    [TestCase("D5", "Other!D5", "D5")] // Delete cell with same coordinate on different sheet.
    [TestCase("D5", "Sheet!D5", "#REF!")] // Fully delete a cell.
    [TestCase("$D$5", "Sheet!D5", "#REF!")] // Only posıtıon matter, reference style doesn't.

    // Row span tests
    [TestCase("2:$4", "Sheet!C:F", "2:$4")] // Row span is basically never changed.
    [TestCase("$2:4", "Sheet!A1:D5", "$2:4")] // Row span is basically never changed.

    // Deleted area is fully upward of the reference.
    [TestCase("B3:D7", "Sheet!B1:D2", "B3:D7")]

    // Deleted area is fully below of the reference.
    [TestCase("B3:D7", "Sheet!B8:D9", "B3:D7")]

    // Deleted area is inside and only partially covers rows of a reference.
    [TestCase("B3:D9", "Sheet!B4:D6", "B3:D9")]

    // Deleted area covers full height of the reference
    [TestCase("A1:B3", "Sheet!D1:E3", "A1:B3")] // Deleted area is completely to the right of the reference.
    [TestCase("E1:I2", "Sheet!I1:K5", "E1:H2")]
    [TestCase("E1:I2", "Sheet!I1:I2", "E1:H2")]
    [TestCase("E1:I2", "Sheet!H1:I2", "E1:G2")]
    [TestCase("E1:I2", "Sheet!E1:I2", "#REF!")]
    [TestCase("E1:I2", "Sheet!E1:F2", "E1:G2")]
    [TestCase("E1:I2", "Sheet!D1:F2", "D1:F2")]
    [TestCase("E1:I2", "Sheet!B1:D2", "B1:F2")]

    // Deleted area covers a top slice of a reference
    [TestCase("B3:D7", "Sheet!B3:D5", "B6:D7")]
    [TestCase("B3:D7", "Sheet!B1:D5", "B6:D7")]
    [TestCase("B3:D7", "Sheet!A1:G4", "B5:D7")]
    [TestCase("B3:D7", "Sheet!C2:D5", "B3:D7")]
    [TestCase("B3:D7", "Sheet!B1:C3", "B3:D7")]
    [TestCase("B3:D7", "Sheet!A2:C5", "B3:D7")]
    [TestCase("B3:D7", "Sheet!E3:F6", "B3:D7")]

    // Deleted area covers bottom slice of a reference
    [TestCase("C5:E12", "Sheet!C7:E12", "C5:E6")]
    [TestCase("C5:E12", "Sheet!C10:E16", "C5:E9")]
    [TestCase("C5:E12", "Sheet!A10:E13", "C5:E9")]
    [TestCase("C5:E12", "Sheet!D8:E13", "C5:E12")]
    [TestCase("C5:E12", "Sheet!C8:D13", "C5:E12")]
    [TestCase("C5:E12", "Sheet!A8:B12", "C5:E12")]
    [TestCase("C5:E12", "Sheet!F8:F12", "C5:E12")]
    public void Delete_area_and_shift_left_reference(string formula, string deletedArea, string expected)
    {
        // TODO: Once incorporated into cell deletion, replace with a public API test case through SUM(reference) in a cell.
        Assert.True(ReferenceParser.TryParseSheetA1(deletedArea, out var deletedSheet, out var deletedReference));
        var deleted = new XLBookArea(deletedSheet, deletedReference.ToSheetRangeA1());

        var result = FormulaConverter.ModifyA1(formula, "Sheet", 1, 1, new ReferenceShiftOnDeleteRefModVisitor(deleted, XLShiftDeletedCells.ShiftCellsLeft));

        Assert.AreEqual(expected, result);
    }
}
