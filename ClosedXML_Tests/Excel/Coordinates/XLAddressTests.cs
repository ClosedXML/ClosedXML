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

        [Test]
        public void EqualityTests()
        {
            // Valid addresses
            var ws = new XLWorkbook().AddWorksheet();
            var validAddress1 = ws.Cell("A1").Address;
            var validAddress2 = ws.Cell("C3").Address;
            var validAddress3 = ((XLRangeAddress)ws.Range("B2:A1").RangeAddress).Normalize().FirstAddress;

            Assert.IsFalse(validAddress1.Equals(validAddress2));
            Assert.IsTrue(validAddress1.Equals(validAddress3));

            // Invalid addresses
            var invalidAddress1 = ProduceInvalidAddress();
            var invalidAddress2 = ProduceAddressOnDeletedWorksheet();
            var invalidAddress3 = ProduceInvalidAddressOnDeletedWorksheet();

            var invalidAddress4 = new XLAddress(invalidAddress1.Worksheet as XLWorksheet, -50, -50, false, false);

            Assert.IsFalse(validAddress1.Equals(invalidAddress2));

            Assert.IsFalse(invalidAddress1.Equals(invalidAddress2));
            Assert.IsFalse(invalidAddress1.Equals(invalidAddress3));
            Assert.IsTrue(invalidAddress1.Equals(invalidAddress4));

            Assert.IsFalse(invalidAddress2.Equals(invalidAddress3));
            Assert.IsFalse(invalidAddress2.Equals(invalidAddress4));

            Assert.IsFalse(invalidAddress3.Equals(invalidAddress4));
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
