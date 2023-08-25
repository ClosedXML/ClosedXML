using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Coordinates
{
    [TestFixture]
    public class XLNameTests
    {
        [Test]
        public void Workbook_scoped_name_is_compared_case_insensitive()
        {
            var lowerCase = new XLName("name");
            var upperCase = new XLName("NAME");

            Assert.AreEqual(lowerCase, upperCase);
            Assert.AreEqual(lowerCase.GetHashCode(), upperCase.GetHashCode());

            Assert.AreNotEqual(lowerCase, new XLName("different_name"));
        }

        [Test]
        public void Sheet_scoped_name_is_compared_case_insensitive()
        {
            var lowerCase = new XLName("sheet", "name");
            var upperCase = new XLName("SHEET", "NAME");

            Assert.AreEqual(lowerCase, upperCase);
            Assert.AreEqual(lowerCase.GetHashCode(), upperCase.GetHashCode());

            Assert.AreNotEqual(lowerCase, new XLName("Different sheet", "name"));
            Assert.AreNotEqual(lowerCase, new XLName("sheet", "different_name"));
        }
    }
}
