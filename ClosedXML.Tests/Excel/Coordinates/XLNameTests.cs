using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Coordinates
{
    [TestFixture]
    public class XLNameTests
    {
        [Test]
        public void Workbook_name_is_compared_case_insensitive()
        {
            var left = new XLName("name");
            var right = new XLName("NAME");

            Assert.AreEqual(left, right);
            Assert.AreEqual(left.GetHashCode(), right.GetHashCode());

            Assert.AreNotEqual(left, new XLName("different_name"));
        }

        [Test]
        public void Sheet_name_is_compared_case_insensitive()
        {
            var left = new XLName("sheet", "name");
            var right = new XLName("SHEET", "NAME");

            Assert.AreEqual(left, right);
            Assert.AreEqual(left.GetHashCode(), right.GetHashCode());

            Assert.AreNotEqual(left, new XLName("Different sheet", "name"));
            Assert.AreNotEqual(left, new XLName("sheet", "different_name"));
        }
    }
}
