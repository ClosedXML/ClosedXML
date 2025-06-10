using System;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Coordinates
{
    [TestFixture]
    public class XLSheetRangeTests
    {
        [TestCase("A1", 1, 1, 1, 1)]
        [TestCase("A1:Z100", 1, 1, 100, 26)]
        [TestCase("BD14:EG256", 14, 56, 256, 137)]
        [TestCase("A1:XFD1048576", 1, 1, 1048576, 16384)]
        [TestCase("XFD1048576", 1048576, 16384, 1048576, 16384)]
        [TestCase("XFD1048576:XFD1048576", 1048576, 16384, 1048576, 16384)]
        public void ParseCellRefsAccordingToGrammar(string refText, int firstRow, int firstCol, int lastRow, int lastCol)
        {
            var reference = XLSheetRange.Parse(refText);
            Assert.AreEqual(firstRow, reference.FirstPoint.Row);
            Assert.AreEqual(firstCol, reference.FirstPoint.Column);
            Assert.AreEqual(lastRow, reference.LastPoint.Row);
            Assert.AreEqual(lastCol, reference.LastPoint.Column);
        }

        [TestCase("")]
        [TestCase("A1:")]
        [TestCase(":A1")]
        [TestCase("A1: A1")]
        [TestCase(" A1:A1")]
        [TestCase("A1:A1 ")]
        [TestCase("B1:A1")]
        [TestCase("A2:A1")]
        public void InvalidInputsAreNotParsed(string invalidRef)
        {
            Assert.Throws<FormatException>(() => XLSheetRange.Parse(invalidRef));
        }

        [TestCase("A1:A1", "A1")]
        [TestCase("DO974:LAR2487", "DO974:LAR2487")]
        [TestCase("XFD1048576:XFD1048576", "XFD1048576")]
        [TestCase("XFD1048575:XFD1048576", "XFD1048575:XFD1048576")]
        public void CanFormatToString(string cellRef, string expected)
        {
            var r = XLSheetRange.Parse(cellRef);
            Assert.AreEqual(expected, r.ToString());
        }

        [TestCase("A1", "A1", "A1")]
        [TestCase("A1", "B3", "A1:B3")]
        [TestCase("C2", "B3", "B2:C3")]
        [TestCase("I6:J9", "L7", "I6:L9")]
        [TestCase("B2:B4", "A3:C3", "A2:C4")]
        [TestCase("B2:C3", "E5:F6", "B2:F6")]
        public void RangeOperation(string leftOperand, string rightOperand, string expectedRange)
        {
            var left = XLSheetRange.Parse(leftOperand);
            var right = XLSheetRange.Parse(rightOperand);
            var expected = XLSheetRange.Parse(expectedRange);

            Assert.AreEqual(expected, left.Range(right));
        }

        [TestCase("A1", "A1", "A1")]
        [TestCase("A1", "A2", null)]
        [TestCase("B1:B3", "A2:C2", "B2")]
        [TestCase("A1:A3", "B2:C2", null)]
        [TestCase("A1:D6", "B2:C3", "B2:C3")]
        [TestCase("A1:C6", "B4:E10", "B4:C6")]
        public void IntersectOperation(string leftOperand, string rightOperand, string expectedRange)
        {
            var left = XLSheetRange.Parse(leftOperand);
            var right = XLSheetRange.Parse(rightOperand);
            var expected = expectedRange is null ? (XLSheetRange?)null : XLSheetRange.Parse(expectedRange);

            Assert.AreEqual(expected, left.Intersect(right));
        }

        [TestCase("A1", "A1", true)]
        [TestCase("A1", "A2", false)]
        [TestCase("B1:B3", "A2:C2", true)]
        [TestCase("A1:A3", "B2:C2", false)]
        [TestCase("A1:D6", "B2:C3", true)]
        [TestCase("A1:C6", "B4:E10", true)]
        public void Intersects_checks_whether_the_range_has_intersection_with_another(string leftOperand, string rightOperand, bool expected)
        {
            var left = XLSheetRange.Parse(leftOperand);
            var right = XLSheetRange.Parse(rightOperand);

            Assert.AreEqual(expected, left.Intersects(right));
        }

        [TestCase("A1", "A1", true)]
        [TestCase("B1:C3", "B1:C3", true)]
        [TestCase("A1:D4", "B2:C3", true)]
        [TestCase("B3:C3", "B2:C3", false)]
        [TestCase("A2:C2", "B2:C3", false)]
        public void Overlaps_checks_whether_left_fully_overlaps_right(string leftOperand, string rightOperand, bool expected)
        {
            var left = XLSheetRange.Parse(leftOperand);
            var right = XLSheetRange.Parse(rightOperand);

            Assert.AreEqual(expected, left.Overlaps(right));
        }

        [TestCase("C4:F8", "C1:F3", "C4:F8")] // Inserted area is fully above
        [TestCase("C4:F8", "A9:G12", "C4:F8")] // Inserted area is fully below
        [TestCase("C4:F8", "G1:H5", "C4:F8")] // Inserted are is fully to the right
        [TestCase("C4:F8", "C1:D11", "E4:H8")] // Inserted area at the left column of the area
        [TestCase("C4:F8", "A1:B8", "E4:H8")] // Inserted area is fully to the left
        [TestCase("C4:F8", "D4:E8", "C4:H8")] // Inserted into the area
        [TestCase("C4:F8", "D2:I8", "C4:L8")] // Inside the area, overlapping = extend
        [TestCase("C4:F8", "F4:F8", "C4:G8")] // Last column of the area, overlapping = extend
        [TestCase("XFD1", "XFB1", null)] // Completely pushed out of the range
        [TestCase("XFA1:XFD1", "XEZ1:XFA1", "XFC1:XFD1")] // Partially pushed out of the range
        [TestCase("XFA1:XFD1", "XFB1:XFC1", "XFA1:XFD1")] // Extend below last row
        public void TryInsertAreaAndShiftRight_without_partial_cover(string original, string inserted, string repositioned)
        {
            var originalArea = XLSheetRange.Parse(original);
            var insertedArea = XLSheetRange.Parse(inserted);
            var repositionedArea = repositioned is not null ? XLSheetRange.Parse(repositioned) : (XLSheetRange?)null;

            var success = originalArea.TryInsertAreaAndShiftRight(insertedArea, out var result);

            Assert.True(success);
            Assert.AreEqual(repositionedArea, result);
        }

        [TestCase("C4:F8", "B3:B4")] // Partially above
        [TestCase("C4:F8", "B5:C7")] // In the middle
        [TestCase("C4:F8", "A5:B9")] // Partially below
        public void TryInsertAreaAndShiftRight_with_partial_cover(string original, string inserted)
        {
            var originalArea = XLSheetRange.Parse(original);
            var insertedArea = XLSheetRange.Parse(inserted);

            Assert.False(originalArea.TryInsertAreaAndShiftRight(insertedArea, out _));
        }

        [TestCase("D6:G10", "A1:C15", "D6:G10")] // Inserted are is fully to the left
        [TestCase("D6:G10", "H1:K15", "D6:G10")] // Inserted are is fully to the right
        [TestCase("D6:G10", "A11:K15", "D6:G10")] // Inserted are is fully below
        [TestCase("D6:G10", "D6:G11", "D12:G16")] // Inserted area at the top row of the area
        [TestCase("D6:G10", "C4:H7", "D10:G14")] // Inserted above the area
        [TestCase("D6:G10", "D7:G9", "D6:G13")] // Inserted into the area
        [TestCase("D6:G10", "A7:H9", "D6:G13")] // Inside the area, overlapping = extend
        [TestCase("D6:G10", "D10:G11", "D6:G12")] // Last row of the area, overlapping = extend
        [TestCase("A1048576", "A1048575", null)] // Completely pushed out of the range
        [TestCase("A1048574:A1048576", "A1048570:A1048571", "A1048576")] // Partially pushed out of the range
        [TestCase("A1048570:A1048572", "A1048571:A1048576", "A1048570:A1048576")] // Extend below last row
        public void TryInsertAreaAndShiftDown_without_partial_cover(string original, string inserted, string repositioned)
        {
            var originalArea = XLSheetRange.Parse(original);
            var insertedArea = XLSheetRange.Parse(inserted);
            var repositionedArea = repositioned is not null ? XLSheetRange.Parse(repositioned) : (XLSheetRange?)null;

            var success = originalArea.TryInsertAreaAndShiftDown(insertedArea, out var result);

            Assert.True(success);
            Assert.AreEqual(repositionedArea, result);
        }

        [TestCase("D6:G10", "A6:E6")] // Left
        [TestCase("D6:G10", "D5:D5")] // Above
        [TestCase("D6:G10", "E7:H15")] // Right
        public void TryInsertAreaAndShiftDown_with_partial_cover(string original, string inserted)
        {
            var originalArea = XLSheetRange.Parse(original);
            var insertedArea = XLSheetRange.Parse(inserted);

            Assert.False(originalArea.TryInsertAreaAndShiftDown(insertedArea, out _));
        }

        [TestCase("E4:G4", "B3:C5", "C4:E4")] // Deleted area fully to the left with overlapping width
        [TestCase("E4:G4", "A2:D5", "A4:C4")] // The deleted are ends exactly at the column to the left of the area
        [TestCase("E4:G4", "F1:F7", "E4:F4")] // The deleted is fully within the area, but not at left/right column
        [TestCase("E4:G4", "E4:G4", null)] // Delete are exactly covers the area
        [TestCase("E4:G4", "A1:Z9", null)] // Delete fully covers the area
        [TestCase("E4:G4", "H1:K10", "E4:G4")] // The deleted is fully to the right of the area.
        [TestCase("E4:G4", "G3:H5", "E4:F4")] // The deleted partially intersects the area and is to the right.
        [TestCase("D4:E4", "A5:F9", "D4:E4")] // Deleted area is fully downward
        [TestCase("D4:E4", "A1:F3", "D4:E4")] // Deleted area is fully upwards
        [TestCase("D4:E4", "A5:F10", "D4:E4")] // Partial deletion is below -> not affected
        public void TryDeleteAreaAndShiftLeft_without_partial_cover(string original, string deleted, string repositioned)
        {
            var originalArea = XLSheetRange.Parse(original);
            var deletedArea = XLSheetRange.Parse(deleted);
            var repositionedArea = repositioned is not null ? XLSheetRange.Parse(repositioned) : (XLSheetRange?)null;

            var success = originalArea.TryDeleteAreaAndShiftLeft(deletedArea, out var result);

            Assert.True(success);
            Assert.AreEqual(repositionedArea, result);
        }

        [TestCase("D4:E8", "A1:B5")] // Partial left
        [TestCase("D4:E8", "D2:E7")] // Partial inside
        [TestCase("D4:E8", "C4:D6")] // Partial left and inside
        public void TryDeleteAreaAndShiftLeft_with_partial_cover(string original, string deleted)
        {
            var originalArea = XLSheetRange.Parse(original);
            var deletedArea = XLSheetRange.Parse(deleted);
            var success = originalArea.TryDeleteAreaAndShiftLeft(deletedArea, out var result);

            Assert.False(success);
            Assert.Null(result);
        }

        [TestCase("B5:B8", "A1:C3", "B2:B5")] // Deleted area fully above (with a row space) with overlapping width
        [TestCase("B5:B8", "A2:C4", "B2:B5")] // The deleted are ends exactly at the row above the area
        [TestCase("B5:B8", "A6:C7", "B5:B6")] // The deleted is fully within the area, but not at top/bottom row
        [TestCase("B5:B8", "A5:C8", null)] // Delete are exactly covers the area
        [TestCase("B5:B8", "A4:C9", null)] // Delete fully covers the area
        [TestCase("B5:B8", "A9:C10", "B5:B8")] // The deleted is fully below the area.
        [TestCase("B5:B8", "A6:C10", "B5:B5")] // The deleted partially intersects the area and is below.
        [TestCase("B5:B8", "A1:A10", "B5:B8")] // Deleted area is fully on the left
        [TestCase("B5:B8", "C1:C10", "B5:B8")] // Deleted area is fully on the right
        [TestCase("B5:D8", "B9:C10", "B5:D8")] // Partial deletion is below -> not affected
        public void TryDeleteAreaAndShiftUp_without_partial_cover(string leftOperand, string deleted, string expected)
        {
            var originalArea = XLSheetRange.Parse(leftOperand);
            var deletedArea = XLSheetRange.Parse(deleted);
            var expectedResult = expected is not null ? XLSheetRange.Parse(expected) : (XLSheetRange?)null;

            var success = originalArea.TryDeleteAreaAndShiftUp(deletedArea, out var result);

            Assert.True(success);
            Assert.AreEqual(expectedResult, result);
        }

        [TestCase("B5:D8", "A1:B3")] // Partial above
        [TestCase("B5:D8", "C6:D8")] // Partial inside
        [TestCase("B5:D8", "B1:B6")] // Partial above and inside
        public void TryDeleteAreaAndShiftUp_with_partial_cover(string leftOperand, string deleted)
        {
            var originalArea = XLSheetRange.Parse(leftOperand);
            var deletedArea = XLSheetRange.Parse(deleted);
            var success = originalArea.TryDeleteAreaAndShiftUp(deletedArea, out var result);

            Assert.False(success);
            Assert.Null(result);
        }

        // Subtrahend doesn't intersect the minuend
        [TestCase("G10:O24", "C10:F24", "G10:O24")] // Left
        [TestCase("G10:O24", "D5:F9", "G10:O24")] // Left Above
        [TestCase("G10:O24", "G3:O9", "G10:O24")] // Above
        [TestCase("G10:O24", "P6:Q9", "G10:O24")] // Above Right
        [TestCase("G10:O24", "P10:R24", "G10:O24")] // Right
        [TestCase("G10:O24", "P24:Q27", "G10:O24")] // Bottom Right
        [TestCase("G10:O24", "G25:O27", "G10:O24")] // Bottom
        [TestCase("G10:O24", "E25:F29", "G10:O24")] // Bottom Left

        // Subtrahend fully covers the minuend
        [TestCase("G10:O24", "F9:P25", null)]
        [TestCase("G10:O24", "G10:O24", null)]

        // Slice off a left side
        [TestCase("G10:O24", "G10:G24", "H10:O24")]
        [TestCase("G10:O24", "D5:L24", "M10:O24")]

        // Slice off a top side
        [TestCase("G10:O24", "G10:O10", "G11:O24")]
        [TestCase("G10:O24", "D3:R14", "G15:O24")]

        // Slice off a right side
        [TestCase("G10:O24", "O10:O24", "G10:N24")]
        [TestCase("G10:O24", "J4:Q30", "G10:I24")]

        // Slice off a bottom side
        [TestCase("G10:O24", "G24:O24", "G10:O23")]
        [TestCase("G10:O24", "E12:P36", "G10:O11")]
        public void TrySubtract_subtracts_one_area_from_another_and_returns_true_when_there_is_no_split(string minuend, string subtrahend, string expected)
        {
            var minuendArea = XLSheetRange.Parse(minuend);
            var subtrahendArea = XLSheetRange.Parse(subtrahend);
            var expectedResult = expected is not null ? XLSheetRange.Parse(expected) : (XLSheetRange?)null;

            var noSplit = minuendArea.TrySubtract(subtrahendArea, out var resultArea);

            Assert.True(noSplit);
            Assert.AreEqual(expectedResult, resultArea);
        }

        [TestCase("G10:O24", "H11:N23")] // Subtrahend is inside the minuend, the result would in in a shape of O
        [TestCase("G10:O24", "G13:I15")] // Result is a horseshoe opened to the left
        [TestCase("G10:O24", "H7:K14")] // Result is a horseshoe opened to the top
        [TestCase("G10:O24", "M13:Q16")] // Result is a horseshoe opened to the right
        [TestCase("G10:O24", "J21:L27")] // Result is a horseshoe opened to the bottom
        [TestCase("G10:O24", "E8:I12")] // Subtract top left corner
        [TestCase("G10:O24", "N4:P12")] // Subtract top right corner
        [TestCase("G10:O24", "K15:P26")] // Subtract bottom right corner
        [TestCase("G10:O24", "E19:H28")] // Subtract bottom left corner
        public void TrySubtract_returns_false_on_split(string minuend, string subtrahend)
        {
            var minuendArea = XLSheetRange.Parse(minuend);
            var subtrahendArea = XLSheetRange.Parse(subtrahend);

            var noSplit = minuendArea.TrySubtract(subtrahendArea, out var resultArea);

            Assert.False(noSplit);
            Assert.IsNull(resultArea);
        }
    }
}
