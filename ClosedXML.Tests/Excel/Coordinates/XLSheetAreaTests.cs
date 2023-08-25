using System.Collections.Generic;
using System.Runtime.Serialization;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Coordinates
{
    [TestFixture]
    public class XLSheetAreaTests
    {
        [Test]
        public void Equality_uses_case_insensitive_comparison_for_sheet_name()
        {
            var upperCase = new XLSheetArea("NAME", new XLSheetRange(1, 2, 3, 4));
            var lowerCase = new XLSheetArea("name", new XLSheetRange(1, 2, 3, 4));
            Assert.AreEqual(upperCase.GetHashCode(), lowerCase.GetHashCode());
            Assert.AreEqual(upperCase, lowerCase);
        }

        [Test]
        public void Intersection_produces_range_intersection_in_same_sheet()
        {
            var left = new XLSheetArea("THIS", XLSheetRange.Parse("A1:C3"));
            var rightSameSheet = new XLSheetArea("this", XLSheetRange.Parse("B2:D4"));
            var rightDifferentSheet = new XLSheetArea("Different", XLSheetRange.Parse("B2:D4"));

            var sameSheetIntersection = left.Intersect(rightSameSheet);
            Assert.AreEqual(new XLSheetArea("this", XLSheetRange.Parse("B2:C3")), sameSheetIntersection);

            var differentSheetIntersection = left.Intersect(rightDifferentSheet);
            Assert.Null(differentSheetIntersection);
        }
    }
}
