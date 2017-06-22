using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class XLAddressTests
    {
        [Test]
        public void ToStringTest()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLAddress address = ws.Cell(1, 1).Address;

            Assert.AreEqual("A1", address.ToString());
            Assert.AreEqual("A1", address.ToString(XLReferenceStyle.A1));
            Assert.AreEqual("R1C1", address.ToString(XLReferenceStyle.R1C1));
            Assert.AreEqual("A1", address.ToString(XLReferenceStyle.Default));

            Assert.AreEqual("A1", address.ToStringRelative());
            Assert.AreEqual("Sheet1!A1", address.ToStringRelative(true));

            Assert.AreEqual("$A$1", address.ToStringFixed());
            Assert.AreEqual("$A$1", address.ToStringFixed(XLReferenceStyle.A1));
            Assert.AreEqual("R1C1", address.ToStringFixed(XLReferenceStyle.R1C1));
            Assert.AreEqual("$A$1", address.ToStringFixed(XLReferenceStyle.Default));
            Assert.AreEqual("Sheet1!$A$1", address.ToStringFixed(XLReferenceStyle.A1, true));
            Assert.AreEqual("Sheet1!R1C1", address.ToStringFixed(XLReferenceStyle.R1C1, true));
            Assert.AreEqual("Sheet1!$A$1", address.ToStringFixed(XLReferenceStyle.Default, true));
        }

        [Test]
        public void ToStringTestWithSpace()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet 1");
            IXLAddress address = ws.Cell(1, 1).Address;

            Assert.AreEqual("A1", address.ToString());
            Assert.AreEqual("A1", address.ToString(XLReferenceStyle.A1));
            Assert.AreEqual("R1C1", address.ToString(XLReferenceStyle.R1C1));
            Assert.AreEqual("A1", address.ToString(XLReferenceStyle.Default));

            Assert.AreEqual("A1", address.ToStringRelative());
            Assert.AreEqual("'Sheet 1'!A1", address.ToStringRelative(true));

            Assert.AreEqual("$A$1", address.ToStringFixed());
            Assert.AreEqual("$A$1", address.ToStringFixed(XLReferenceStyle.A1));
            Assert.AreEqual("R1C1", address.ToStringFixed(XLReferenceStyle.R1C1));
            Assert.AreEqual("$A$1", address.ToStringFixed(XLReferenceStyle.Default));
            Assert.AreEqual("'Sheet 1'!$A$1", address.ToStringFixed(XLReferenceStyle.A1, true));
            Assert.AreEqual("'Sheet 1'!R1C1", address.ToStringFixed(XLReferenceStyle.R1C1, true));
            Assert.AreEqual("'Sheet 1'!$A$1", address.ToStringFixed(XLReferenceStyle.Default, true));
        }
    }
}
