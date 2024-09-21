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
    }
}
