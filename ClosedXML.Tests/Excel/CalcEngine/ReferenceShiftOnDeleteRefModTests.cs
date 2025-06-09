using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine.Visitors;
using ClosedXML.Extensions;
using ClosedXML.Parser;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine;

[TestFixture]
[TestOf(typeof(ReferenceShiftOnDeleteRefModVisitor))]
public class ReferenceShiftOnDeleteRefModTests
{
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
    [TestCase("A1:D3", "Sheet!A4:D4", "A1:D3")] // Deleted area is completely below the reference.
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
}
