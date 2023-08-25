using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Coordinates
{
    [TestFixture]
    public class XLSheetAreaTests
    {
        [Test]
        public void Sheet_name_is_compared_case_insensitive()
        {
            var upperCase = new XLSheetArea("NAME", new XLSheetRange(1, 2, 3, 4));
            var lowerCase = new XLSheetArea("name", new XLSheetRange(1, 2, 3, 4));
            Assert.AreEqual(upperCase.GetHashCode(), lowerCase.GetHashCode());
            Assert.AreEqual(upperCase, lowerCase);
        }

        [Test]
        public void Intersection_produces_range_intersection_in_same_sheet()
        {
            var sheetArea1 = new XLSheetArea("SHEET", XLSheetRange.Parse("A1:C3"));
            var sheetArea2 = new XLSheetArea("sheet", XLSheetRange.Parse("B2:D4"));
            var otherSheetArea = new XLSheetArea("Other", XLSheetRange.Parse("B2:D4"));

            var sameSheetIntersection = sheetArea1.Intersect(sheetArea2);
            Assert.AreEqual(new XLSheetArea("sheet", XLSheetRange.Parse("B2:C3")), sameSheetIntersection);

            var differentSheetIntersection = sheetArea1.Intersect(otherSheetArea);
            Assert.Null(differentSheetIntersection);
        }
    }
}
