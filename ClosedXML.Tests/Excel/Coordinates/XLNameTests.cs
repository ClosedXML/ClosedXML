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
            
            Assert.Multiple(() =>
            {
                Assert.That(upperCase, Is.EqualTo(lowerCase));
                Assert.That(upperCase.GetHashCode(), Is.EqualTo(lowerCase.GetHashCode()));
                Assert.That(new XLName("different_name"), Is.Not.EqualTo(lowerCase));
            });
        }

        [Test]
        public void Sheet_scoped_name_is_compared_case_insensitive()
        {
            var lowerCase = new XLName("sheet", "name");
            var upperCase = new XLName("SHEET", "NAME");

            
            Assert.Multiple(() =>
            {
                Assert.That(upperCase, Is.EqualTo(lowerCase));
                Assert.That(upperCase.GetHashCode(), Is.EqualTo(lowerCase.GetHashCode()));
                Assert.That(new XLName("Different sheet", "name"), Is.Not.EqualTo(lowerCase));
                Assert.That(new XLName("sheet", "different_name"), Is.Not.EqualTo(lowerCase));
            });
        }
    }
}
