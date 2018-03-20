using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class XLRangeAddressTests
    {
        [Test]
        public void ToStringTest()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRangeAddress address = ws.Cell(1, 1).AsRange().RangeAddress;

            Assert.AreEqual("A1:A1", address.ToString());
            Assert.AreEqual("Sheet1!R1C1:R1C1", address.ToString(XLReferenceStyle.R1C1, true));

            Assert.AreEqual("A1:A1", address.ToStringRelative());
            Assert.AreEqual("Sheet1!A1:A1", address.ToStringRelative(true));

            Assert.AreEqual("$A$1:$A$1", address.ToStringFixed());
            Assert.AreEqual("$A$1:$A$1", address.ToStringFixed(XLReferenceStyle.A1));
            Assert.AreEqual("R1C1:R1C1", address.ToStringFixed(XLReferenceStyle.R1C1));
            Assert.AreEqual("$A$1:$A$1", address.ToStringFixed(XLReferenceStyle.Default));
            Assert.AreEqual("Sheet1!$A$1:$A$1", address.ToStringFixed(XLReferenceStyle.A1, true));
            Assert.AreEqual("Sheet1!R1C1:R1C1", address.ToStringFixed(XLReferenceStyle.R1C1, true));
            Assert.AreEqual("Sheet1!$A$1:$A$1", address.ToStringFixed(XLReferenceStyle.Default, true));
        }

        [Test]
        public void ToStringTestWithSpace()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet 1");
            IXLRangeAddress address = ws.Cell(1, 1).AsRange().RangeAddress;

            Assert.AreEqual("A1:A1", address.ToString());
            Assert.AreEqual("'Sheet 1'!R1C1:R1C1", address.ToString(XLReferenceStyle.R1C1, true));

            Assert.AreEqual("A1:A1", address.ToStringRelative());
            Assert.AreEqual("'Sheet 1'!A1:A1", address.ToStringRelative(true));

            Assert.AreEqual("$A$1:$A$1", address.ToStringFixed());
            Assert.AreEqual("$A$1:$A$1", address.ToStringFixed(XLReferenceStyle.A1));
            Assert.AreEqual("R1C1:R1C1", address.ToStringFixed(XLReferenceStyle.R1C1));
            Assert.AreEqual("$A$1:$A$1", address.ToStringFixed(XLReferenceStyle.Default));
            Assert.AreEqual("'Sheet 1'!$A$1:$A$1", address.ToStringFixed(XLReferenceStyle.A1, true));
            Assert.AreEqual("'Sheet 1'!R1C1:R1C1", address.ToStringFixed(XLReferenceStyle.R1C1, true));
            Assert.AreEqual("'Sheet 1'!$A$1:$A$1", address.ToStringFixed(XLReferenceStyle.Default, true));
        }

        [TestCase("B2:E5", "B2:E5")]
        [TestCase("E5:B2", "B2:E5")]
        [TestCase("B5:E2", "B2:E5")]
        [TestCase("B2:E$5", "B2:E$5")]
        [TestCase("B2:$E$5", "B2:$E$5")]
        [TestCase("B$2:$E$5", "B$2:$E$5")]
        [TestCase("$B$2:$E$5", "$B$2:$E$5")]
        [TestCase("B5:E$2", "B$2:E5")]
        [TestCase("$B$5:E2", "$B2:E$5")]
        [TestCase("$B$5:E$2", "$B$2:E$5")]
        [TestCase("$B$5:$E$2", "$B$2:$E$5")]
        public void RangeAddressNormalizeTest(string inputAddress, string expectedAddress)
        {
            XLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet 1") as XLWorksheet;
            var rangeAddress = new XLRangeAddress(ws, inputAddress);

            var normalizedAddress = rangeAddress.Normalize();

            Assert.ReferenceEquals(ws, rangeAddress.Worksheet);
            Assert.AreEqual(expectedAddress, normalizedAddress.ToString());
        }
    }
}
