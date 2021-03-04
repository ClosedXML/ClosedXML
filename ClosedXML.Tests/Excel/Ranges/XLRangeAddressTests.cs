using ClosedXML.Excel;
using NUnit.Framework;
using System;

namespace ClosedXML.Tests
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

            Assert.AreSame(ws, rangeAddress.Worksheet);
            Assert.AreEqual(expectedAddress, normalizedAddress.ToString());
        }

        [Test]
        public void InvalidRangeAddressToStringTest()
        {
            var address = ProduceInvalidAddress();

            Assert.AreEqual("#REF!", address.ToString());
            Assert.AreEqual("#REF!", address.ToString(XLReferenceStyle.A1));
            Assert.AreEqual("#REF!", address.ToString(XLReferenceStyle.Default));
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToString(XLReferenceStyle.R1C1));
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToString(XLReferenceStyle.A1, true));
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToString(XLReferenceStyle.Default, true));
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToString(XLReferenceStyle.R1C1, true));
        }

        [Test]
        public void InvalidRangeAddressToStringFixedTest()
        {
            var address = ProduceInvalidAddress();

            Assert.AreEqual("#REF!", address.ToStringFixed());
            Assert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.A1));
            Assert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.Default));
            Assert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.R1C1));
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToStringFixed(XLReferenceStyle.A1, true));
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToStringFixed(XLReferenceStyle.Default, true));
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToStringFixed(XLReferenceStyle.R1C1, true));
        }

        [Test]
        public void InvalidRangeAddressToStringRelativeTest()
        {
            var address = ProduceInvalidAddress();

            Assert.AreEqual("#REF!", address.ToStringRelative());
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToStringRelative(true));
        }

        [Test]
        public void RangeAddressOnDeletedWorksheetToStringTest()
        {
            var address = ProduceAddressOnDeletedWorksheet();

            Assert.AreEqual("#REF!A1:B2", address.ToString());
            Assert.AreEqual("#REF!A1:B2", address.ToString(XLReferenceStyle.A1));
            Assert.AreEqual("#REF!A1:B2", address.ToString(XLReferenceStyle.Default));
            Assert.AreEqual("#REF!R1C1:R2C2", address.ToString(XLReferenceStyle.R1C1));
            Assert.AreEqual("#REF!A1:B2", address.ToString(XLReferenceStyle.A1, true));
            Assert.AreEqual("#REF!A1:B2", address.ToString(XLReferenceStyle.Default, true));
            Assert.AreEqual("#REF!R1C1:R2C2", address.ToString(XLReferenceStyle.R1C1, true));
        }

        [Test]
        public void RangeAddressOnDeletedWorksheetToStringFixedTest()
        {
            var address = ProduceAddressOnDeletedWorksheet();

            Assert.AreEqual("#REF!$A$1:$B$2", address.ToStringFixed());
            Assert.AreEqual("#REF!$A$1:$B$2", address.ToStringFixed(XLReferenceStyle.A1));
            Assert.AreEqual("#REF!$A$1:$B$2", address.ToStringFixed(XLReferenceStyle.Default));
            Assert.AreEqual("#REF!R1C1:R2C2", address.ToStringFixed(XLReferenceStyle.R1C1));
            Assert.AreEqual("#REF!$A$1:$B$2", address.ToStringFixed(XLReferenceStyle.A1, true));
            Assert.AreEqual("#REF!$A$1:$B$2", address.ToStringFixed(XLReferenceStyle.Default, true));
            Assert.AreEqual("#REF!R1C1:R2C2", address.ToStringFixed(XLReferenceStyle.R1C1, true));
        }

        [Test]
        public void RangeAddressOnDeletedWorksheetToStringRelativeTest()
        {
            var address = ProduceAddressOnDeletedWorksheet();

            Assert.AreEqual("#REF!A1:B2", address.ToStringRelative());
            Assert.AreEqual("#REF!A1:B2", address.ToStringRelative(true));
        }

        [Test]
        public void InvalidRangeAddressOnDeletedWorksheetToStringTest()
        {
            var address = ProduceInvalidAddressOnDeletedWorksheet();

            Assert.AreEqual("#REF!#REF!", address.ToString());
            Assert.AreEqual("#REF!#REF!", address.ToString(XLReferenceStyle.A1));
            Assert.AreEqual("#REF!#REF!", address.ToString(XLReferenceStyle.Default));
            Assert.AreEqual("#REF!#REF!", address.ToString(XLReferenceStyle.R1C1));
            Assert.AreEqual("#REF!#REF!", address.ToString(XLReferenceStyle.A1, true));
            Assert.AreEqual("#REF!#REF!", address.ToString(XLReferenceStyle.Default, true));
            Assert.AreEqual("#REF!#REF!", address.ToString(XLReferenceStyle.R1C1, true));
        }

        [Test]
        public void InvalidRangeAddressOnDeletedWorksheetToStringFixedTest()
        {
            var address = ProduceInvalidAddressOnDeletedWorksheet();

            Assert.AreEqual("#REF!#REF!", address.ToStringFixed());
            Assert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.A1));
            Assert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.Default));
            Assert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.R1C1));
            Assert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.A1, true));
            Assert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.Default, true));
            Assert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.R1C1, true));
        }

        [Test]
        public void InvalidRangeAddressOnDeletedWorksheetToStringRelativeTest()
        {
            var address = ProduceInvalidAddressOnDeletedWorksheet();

            Assert.AreEqual("#REF!#REF!", address.ToStringRelative());
            Assert.AreEqual("#REF!#REF!", address.ToStringRelative(true));
        }

        [Test]
        public void FullSpanAddressCannotChange()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                var wsRange = ws.AsRange();
                var row = ws.FirstRow().RowBelow(4).AsRange();
                var column = ws.FirstColumn().ColumnRight(4).AsRange();

                Assert.AreEqual($"1:{XLHelper.MaxRowNumber}", wsRange.RangeAddress.ToString());
                Assert.AreEqual("5:5", row.RangeAddress.ToString());
                Assert.AreEqual("E:E", column.RangeAddress.ToString());

                ws.Columns("Y:Z").Delete();
                ws.Rows("9:10").Delete();

                Assert.AreEqual($"1:{XLHelper.MaxRowNumber}", wsRange.RangeAddress.ToString());
                Assert.AreEqual("5:5", row.RangeAddress.ToString());
                Assert.AreEqual("E:E", column.RangeAddress.ToString());
            }
        }

        [Test]
        public void RangeAddressIsNormalized()
        {
            var ws = new XLWorkbook().AddWorksheet();

            XLRangeAddress rangeAddress;

            rangeAddress = (XLRangeAddress)ws.Range(ws.Cell("A1"), ws.Cell("C3")).RangeAddress;
            Assert.IsTrue(rangeAddress.IsNormalized);

            rangeAddress = (XLRangeAddress)ws.Range(ws.Cell("C3"), ws.Cell("A1")).RangeAddress;
            Assert.IsFalse(rangeAddress.IsNormalized);

            rangeAddress = (XLRangeAddress)ws.Range("B2:B1").RangeAddress;
            Assert.IsFalse(rangeAddress.IsNormalized);

            rangeAddress = (XLRangeAddress)ws.Range("B2:B10").RangeAddress;
            Assert.IsTrue(rangeAddress.IsNormalized);

            rangeAddress = (XLRangeAddress)ws.Range("B:B").RangeAddress;
            Assert.IsTrue(rangeAddress.IsNormalized);

            rangeAddress = (XLRangeAddress)ws.Range("2:2").RangeAddress;
            Assert.IsTrue(rangeAddress.IsNormalized);

            rangeAddress = (XLRangeAddress)ws.RangeAddress;
            Assert.IsTrue(rangeAddress.IsNormalized);
        }

        [Test]
        public void AsRangeTests()
        {
            XLRangeAddress rangeAddress;
            rangeAddress = new XLRangeAddress
            (
                new XLAddress(1, 1, false, false),
                new XLAddress(5, 5, false, false)
            );

            Assert.IsTrue(rangeAddress.IsValid);
            Assert.IsTrue(rangeAddress.IsNormalized);
            Assert.Throws<InvalidOperationException>(() => rangeAddress.AsRange());

            var ws = new XLWorkbook().AddWorksheet() as XLWorksheet;
            rangeAddress = new XLRangeAddress
            (
                new XLAddress(ws, 1, 1, false, false),
                new XLAddress(ws, 5, 5, false, false)
            );

            Assert.IsTrue(rangeAddress.IsValid);
            Assert.IsTrue(rangeAddress.IsNormalized);
            Assert.DoesNotThrow(() => rangeAddress.AsRange());
        }

        [Test]
        public void RelativeRanges()
        {
            var ws = new XLWorkbook().AddWorksheet();

            IXLRangeAddress rangeAddress;

            rangeAddress = ws.Range("D4:E4").RangeAddress.Relative(ws.Range("A1:E4").RangeAddress, ws.Range("B10:F14").RangeAddress);
            Assert.IsTrue(rangeAddress.IsValid);
            Assert.AreEqual("E13:F13", rangeAddress.ToString());

            rangeAddress = ws.Range("D4:E4").RangeAddress.Relative(ws.Range("B10:F14").RangeAddress, ws.Range("A1:E4").RangeAddress);
            Assert.IsFalse(rangeAddress.IsValid);
            Assert.AreEqual("#REF!", rangeAddress.ToString());

            rangeAddress = ws.Range("C3").RangeAddress.Relative(ws.Range("A1:B2").RangeAddress, ws.Range("C3").RangeAddress);
            Assert.IsTrue(rangeAddress.IsValid);
            Assert.AreEqual("E5:E5", rangeAddress.ToString());

            rangeAddress = ws.Range("B2").RangeAddress.Relative(ws.Range("A1").RangeAddress, ws.Range("C3").RangeAddress);
            Assert.IsTrue(rangeAddress.IsValid);
            Assert.AreEqual("D4:D4", rangeAddress.ToString());

            rangeAddress = ws.Range("A1").RangeAddress.Relative(ws.Range("B2").RangeAddress, ws.Range("A1").RangeAddress);
            Assert.IsFalse(rangeAddress.IsValid);
            Assert.AreEqual("#REF!", rangeAddress.ToString());
        }

        [Test]
        public void TestSpanProperties()
        {
            var ws = new XLWorkbook().AddWorksheet() as XLWorksheet;

            var range = ws.Range("B3:E5");
            var rangeAddress = range.RangeAddress as IXLRangeAddress;
            Assert.AreEqual(4, rangeAddress.ColumnSpan);
            Assert.AreEqual(3, rangeAddress.RowSpan);
            Assert.AreEqual(12, rangeAddress.NumberOfCells);

            range = ws.Range("E5:B3");
            rangeAddress = range.RangeAddress as IXLRangeAddress;
            Assert.AreEqual(4, rangeAddress.ColumnSpan);
            Assert.AreEqual(3, rangeAddress.RowSpan);
            Assert.AreEqual(12, rangeAddress.NumberOfCells);

            rangeAddress = ProduceAddressOnDeletedWorksheet();
            Assert.AreEqual(2, rangeAddress.ColumnSpan);
            Assert.AreEqual(2, rangeAddress.RowSpan);
            Assert.AreEqual(4, rangeAddress.NumberOfCells);

            rangeAddress = ProduceInvalidAddress();
            Assert.Throws<InvalidOperationException>(() => { var x = rangeAddress.ColumnSpan; });
            Assert.Throws<InvalidOperationException>(() => { var x = rangeAddress.RowSpan; });
            Assert.Throws<InvalidOperationException>(() => { var x = rangeAddress.NumberOfCells; });
        }

        #region Private Methods

        private IXLRangeAddress ProduceInvalidAddress()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet 1");
            var range = ws.Range("A1:B2");

            ws.Rows(1, 5).Delete();
            return range.RangeAddress;
        }

        private IXLRangeAddress ProduceAddressOnDeletedWorksheet()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet 1");
            var address = ws.Range("A1:B2").RangeAddress;

            ws.Delete();
            return address;
        }

        private IXLRangeAddress ProduceInvalidAddressOnDeletedWorksheet()
        {
            var address = ProduceInvalidAddress();
            address.Worksheet.Delete();
            return address;
        }

        #endregion Private Methods
    }
}
