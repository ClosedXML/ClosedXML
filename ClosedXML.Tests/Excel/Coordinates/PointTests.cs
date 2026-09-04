using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;

namespace ClosedXML.Tests.Excel.Coordinates
{
    [TestFixture]
    public class PointTests
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
            var sheetPoint = Point.Parse(cellRef.AsSpan());
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
            Assert.Throws<FormatException>(() => Point.Parse(cellRef.AsSpan()));
        }

        [TestCase("A1")]
        [TestCase("DE1")]
        [TestCase("D174")]
        [TestCase("XFD1048576")]
        public void CanFormatToString(string cellRef)
        {
            var r = Point.Parse(cellRef);
            Assert.AreEqual(cellRef, r.ToString());
        }

        [Test]
        [Issue("2881")]
        public void Hash_codes_of_points_in_area_have_few_collisions()
        {
            // Points are used in sets or dictionaries. Make sure the hash function doesn't produce
            // too many collision. The hash function originally produced hash codes only from
            // a small range of values instead of all possible integer values. That significantly
            // increased number of collisions and led to a bad performance.
            var hashes = new HashSet<int>();
            for (var row = 1; row <= 1000; ++row)
            {
                for (var column = 1; column <= 100; ++column)
                    hashes.Add(new Point(row, column).GetHashCode());
            }

            // Some collisions are inevitable, but the vast majority of points must be distinguishable.
            Assert.That(hashes.Count, Is.GreaterThan(99000));
        }
    }
}
