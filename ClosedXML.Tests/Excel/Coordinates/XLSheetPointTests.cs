using ClosedXML.Excel;
using NUnit.Framework;
using System;

namespace ClosedXML.Tests.Excel.Coordinates
{
    [TestFixture]
    public class XLSheetPointTests
    {
        [TestCase("A1", 1, 1)]
        [TestCase("AA1", 27, 1)]
        [TestCase("AAA1", 703, 1)]
        [TestCase("Z1", 26, 1)]
        [TestCase("ZZ1", 702, 1)]
        [TestCase("XFD1", 16384, 1)]
        [TestCase("A1", 1, 1)]
        [TestCase("A999", 1, 999)]
        [TestCase("XFD1048576", 16384, 1048576)]
        public void ParseCellRefsAccordingToGrammar(string cellRef, int columnNumber, int rowNumber)
        {
            var sheetPoint = XLSheetPoint.Parse(cellRef.AsSpan());
            Assert.AreEqual(columnNumber, sheetPoint.Column);
            Assert.AreEqual(rowNumber, sheetPoint.Row);
        }

        [TestCase("")]
        [TestCase(" ")]
        [TestCase("A")]
        [TestCase("AA")]
        [TestCase("1")]
        [TestCase("11")]
        [TestCase(" A1")]
        [TestCase("A1 ")]
        [TestCase("A 1")]
        [TestCase("@1")] // @ is a char 'A' - 1
        [TestCase("[1")] // [ is a char 'Z' + 1
        [TestCase("A:")] // : is a char '9' + 1
        [TestCase("A/")] // / is a char '0' - 1
        [TestCase("A1:")]
        [TestCase("A1/")]
        [TestCase("A@1")]
        [TestCase("A[1")]
        [TestCase("XFE1")]
        [TestCase("AAAA1")]
        [TestCase("A1048577")]
        [TestCase("A01")]
        [TestCase("A0")]
        [TestCase("A-1")]
        public void InvalidInputsAreNotParsed(string cellRef)
        {
            Assert.Throws<FormatException>(() => XLSheetPoint.Parse(cellRef.AsSpan()));
        }

        [TestCase("A1")]
        [TestCase("DE1")]
        [TestCase("D174")]
        [TestCase("XFD1048576")]
        public void CanFormatToString(string cellRef)
        {
            var r = XLSheetPoint.Parse(cellRef);
            Assert.AreEqual(cellRef, r.ToString());
        }
    }
}
