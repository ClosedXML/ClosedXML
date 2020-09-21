using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class XLHelperTests
    {
        [Test]
        public void IsValidColumnTest()
        {
            Assert.AreEqual(false, XLHelper.IsValidColumn(""));
            Assert.AreEqual(false, XLHelper.IsValidColumn("1"));
            Assert.AreEqual(false, XLHelper.IsValidColumn("A1"));
            Assert.AreEqual(false, XLHelper.IsValidColumn("AA1"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("A"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("AA"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("AAA"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("Z"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("ZZ"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("XFD"));
            Assert.AreEqual(false, XLHelper.IsValidColumn("ZAA"));
            Assert.AreEqual(false, XLHelper.IsValidColumn("XZA"));
            Assert.AreEqual(false, XLHelper.IsValidColumn("XFZ"));
        }

        [Test]
        public void ReplaceRelative1()
        {
            string result = XLHelper.ReplaceRelative("A1", 2, "B");
            Assert.AreEqual("B2", result);
        }

        [Test]
        public void ReplaceRelative2()
        {
            string result = XLHelper.ReplaceRelative("$A1", 2, "B");
            Assert.AreEqual("$A2", result);
        }

        [Test]
        public void ReplaceRelative3()
        {
            string result = XLHelper.ReplaceRelative("A$1", 2, "B");
            Assert.AreEqual("B$1", result);
        }

        [Test]
        public void ReplaceRelative4()
        {
            string result = XLHelper.ReplaceRelative("$A$1", 2, "B");
            Assert.AreEqual("$A$1", result);
        }

        [Test]
        public void ReplaceRelative5()
        {
            string result = XLHelper.ReplaceRelative("1:1", 2, "B");
            Assert.AreEqual("2:2", result);
        }

        [Test]
        public void ReplaceRelative6()
        {
            string result = XLHelper.ReplaceRelative("$1:1", 2, "B");
            Assert.AreEqual("$1:2", result);
        }

        [Test]
        public void ReplaceRelative7()
        {
            string result = XLHelper.ReplaceRelative("1:$1", 2, "B");
            Assert.AreEqual("2:$1", result);
        }

        [Test]
        public void ReplaceRelative8()
        {
            string result = XLHelper.ReplaceRelative("$1:$1", 2, "B");
            Assert.AreEqual("$1:$1", result);
        }

        [Test]
        public void ReplaceRelative9()
        {
            string result = XLHelper.ReplaceRelative("A:A", 2, "B");
            Assert.AreEqual("B:B", result);
        }

        [Test]
        public void ReplaceRelativeA()
        {
            string result = XLHelper.ReplaceRelative("$A:A", 2, "B");
            Assert.AreEqual("$A:B", result);
        }

        [Test]
        public void ReplaceRelativeB()
        {
            string result = XLHelper.ReplaceRelative("A:$A", 2, "B");
            Assert.AreEqual("B:$A", result);
        }

        [Test]
        public void ReplaceRelativeC()
        {
            string result = XLHelper.ReplaceRelative("$A:$A", 2, "B");
            Assert.AreEqual("$A:$A", result);
        }
    }
}
