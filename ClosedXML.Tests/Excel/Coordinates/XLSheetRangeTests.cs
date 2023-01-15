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
    }
}
