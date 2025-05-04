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
            var upperCase = new XLBookArea("NAME", new XLSheetRange(1, 2, 3, 4));
            var lowerCase = new XLBookArea("name", new XLSheetRange(1, 2, 3, 4));
            Assert.Multiple(() =>
            {
                Assert.That(lowerCase.GetHashCode(), Is.EqualTo(upperCase.GetHashCode()));
                Assert.That(lowerCase, Is.EqualTo(upperCase));
            });
        }

        [Test]
        public void Intersection_produces_range_intersection_in_same_sheet()
        {
            var sheetArea1 = new XLBookArea("SHEET", XLSheetRange.Parse("A1:C3"));
            var sheetArea2 = new XLBookArea("sheet", XLSheetRange.Parse("B2:D4"));
            var otherSheetArea = new XLBookArea("Other", XLSheetRange.Parse("B2:D4"));

            var sameSheetIntersection = sheetArea1.Intersect(sheetArea2);
            Assert.That(sameSheetIntersection, Is.EqualTo(new XLBookArea("sheet", XLSheetRange.Parse("B2:C3"))));

            var differentSheetIntersection = sheetArea1.Intersect(otherSheetArea);
            Assert.Null(differentSheetIntersection);
        }
    }
}
