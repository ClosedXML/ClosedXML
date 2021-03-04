using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests
{
    [TestFixture]
    public class XLAddressTests
    {
        [Test]
        public void ToStringTest()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLAddress address = ws.Cell(1, 1).Address;

            Assert.AreEqual("A1", address.ToString());
            Assert.AreEqual("A1", address.ToString(XLReferenceStyle.A1));
            Assert.AreEqual("R1C1", address.ToString(XLReferenceStyle.R1C1));
            Assert.AreEqual("A1", address.ToString(XLReferenceStyle.Default));
            Assert.AreEqual("Sheet1!A1", address.ToString(XLReferenceStyle.Default, true));

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
            Assert.AreEqual("'Sheet 1'!A1", address.ToString(XLReferenceStyle.Default, true));

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

        [Test]
        public void InvalidAddressToStringTest()
        {
            var address = ProduceInvalidAddress();

            Assert.AreEqual("#REF!", address.ToString());
            Assert.AreEqual("#REF!", address.ToString(XLReferenceStyle.A1));
            Assert.AreEqual("#REF!", address.ToString(XLReferenceStyle.R1C1));
            Assert.AreEqual("#REF!", address.ToString(XLReferenceStyle.Default));
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToString(XLReferenceStyle.Default, true));
        }

        [Test]
        public void InvalidAddressToStringFixedTest()
        {
            var address = ProduceInvalidAddress();

            Assert.AreEqual("#REF!", address.ToStringFixed());
            Assert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.A1));
            Assert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.R1C1));
            Assert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.Default));
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToStringFixed(XLReferenceStyle.A1, true));
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToStringFixed(XLReferenceStyle.R1C1, true));
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToStringFixed(XLReferenceStyle.Default, true));
        }

        [Test]
        public void InvalidAddressToStringRelativeTest()
        {
            var address = ProduceInvalidAddress();

            Assert.AreEqual("#REF!", address.ToStringRelative());
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToStringRelative(true));
        }

        [Test]
        public void AddressOnDeletedWorksheetToStringTest()
        {
            var address = ProduceAddressOnDeletedWorksheet();

            Assert.AreEqual("A1", address.ToString());
            Assert.AreEqual("A1", address.ToString(XLReferenceStyle.A1));
            Assert.AreEqual("R1C1", address.ToString(XLReferenceStyle.R1C1));
            Assert.AreEqual("A1", address.ToString(XLReferenceStyle.Default));
            Assert.AreEqual("#REF!A1", address.ToString(XLReferenceStyle.Default, true));
        }

        [Test]
        public void AddressOnDeletedWorksheetToStringFixedTest()
        {
            var address = ProduceAddressOnDeletedWorksheet();

            Assert.AreEqual("$A$1", address.ToStringFixed());
            Assert.AreEqual("$A$1", address.ToStringFixed(XLReferenceStyle.A1));
            Assert.AreEqual("R1C1", address.ToStringFixed(XLReferenceStyle.R1C1));
            Assert.AreEqual("$A$1", address.ToStringFixed(XLReferenceStyle.Default));
            Assert.AreEqual("#REF!$A$1", address.ToStringFixed(XLReferenceStyle.A1, true));
            Assert.AreEqual("#REF!R1C1", address.ToStringFixed(XLReferenceStyle.R1C1, true));
            Assert.AreEqual("#REF!$A$1", address.ToStringFixed(XLReferenceStyle.Default, true));
        }

        [Test]
        public void AddressOnDeletedWorksheetToStringRelativeTest()
        {
            var address = ProduceAddressOnDeletedWorksheet();

            Assert.AreEqual("A1", address.ToStringRelative());
            Assert.AreEqual("#REF!A1", address.ToStringRelative(true));
        }

        [Test]
        public void InvalidAddressOnDeletedWorksheetToStringTest()
        {
            var address = ProduceInvalidAddressOnDeletedWorksheet();

            Assert.AreEqual("#REF!", address.ToString());
            Assert.AreEqual("#REF!", address.ToString(XLReferenceStyle.A1));
            Assert.AreEqual("#REF!", address.ToString(XLReferenceStyle.R1C1));
            Assert.AreEqual("#REF!", address.ToString(XLReferenceStyle.Default));
            Assert.AreEqual("#REF!#REF!", address.ToString(XLReferenceStyle.Default, true));
        }

        [Test]
        public void InvalidAddressOnDeletedWorksheetToStringFixedTest()
        {
            var address = ProduceInvalidAddressOnDeletedWorksheet();

            Assert.AreEqual("#REF!", address.ToStringFixed());
            Assert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.A1));
            Assert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.R1C1));
            Assert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.Default));
            Assert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.A1, true));
            Assert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.R1C1, true));
            Assert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.Default, true));
        }

        [Test]
        public void InvalidAddressOnDeletedWorksheetToStringRelativeTest()
        {
            var address = ProduceInvalidAddressOnDeletedWorksheet();

            Assert.AreEqual("#REF!", address.ToStringRelative());
            Assert.AreEqual("#REF!#REF!", address.ToStringRelative(true));
        }

        #region Private Methods

        private IXLAddress ProduceInvalidAddress()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet 1");
            var range = ws.Range("A1:B2");

            ws.Rows(1, 5).Delete();
            return range.RangeAddress.FirstAddress;
        }

        private IXLAddress ProduceAddressOnDeletedWorksheet()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet 1");
            var address = ws.Cell("A1").Address;

            ws.Delete();
            return address;
        }

        private IXLAddress ProduceInvalidAddressOnDeletedWorksheet()
        {
            var address = ProduceInvalidAddress();
            address.Worksheet.Delete();
            return address;
        }

        #endregion Private Methods
    }
}
