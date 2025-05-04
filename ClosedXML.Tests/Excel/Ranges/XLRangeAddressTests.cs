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

            Assert.Multiple(() =>
            {
                Assert.That(address.ToString(), Is.EqualTo("A1:A1"));
                Assert.That(address.ToString(XLReferenceStyle.R1C1, true), Is.EqualTo("Sheet1!R1C1:R1C1"));

                Assert.That(address.ToStringRelative(), Is.EqualTo("A1:A1"));
                Assert.That(address.ToStringRelative(true), Is.EqualTo("Sheet1!A1:A1"));

                Assert.That(address.ToStringFixed(), Is.EqualTo("$A$1:$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1), Is.EqualTo("$A$1:$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1), Is.EqualTo("R1C1:R1C1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default), Is.EqualTo("$A$1:$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true), Is.EqualTo("Sheet1!$A$1:$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true), Is.EqualTo("Sheet1!R1C1:R1C1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("Sheet1!$A$1:$A$1"));
            });
        }

        [Test]
        public void ToStringTestWithSpace()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet 1");
            IXLRangeAddress address = ws.Cell(1, 1).AsRange().RangeAddress;

            Assert.Multiple(() =>
            {
                Assert.That(address.ToString(), Is.EqualTo("A1:A1"));
                Assert.That(address.ToString(XLReferenceStyle.R1C1, true), Is.EqualTo("'Sheet 1'!R1C1:R1C1"));

                Assert.That(address.ToStringRelative(), Is.EqualTo("A1:A1"));
                Assert.That(address.ToStringRelative(true), Is.EqualTo("'Sheet 1'!A1:A1"));

                Assert.That(address.ToStringFixed(), Is.EqualTo("$A$1:$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1), Is.EqualTo("$A$1:$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1), Is.EqualTo("R1C1:R1C1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default), Is.EqualTo("$A$1:$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!$A$1:$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true), Is.EqualTo("'Sheet 1'!R1C1:R1C1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("'Sheet 1'!$A$1:$A$1"));
            });
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

            Assert.Multiple(() =>
            {
                Assert.That(rangeAddress.Worksheet, Is.SameAs(ws));
                Assert.That(normalizedAddress.ToString(), Is.EqualTo(expectedAddress));
            });
        }

        [Test]
        public void InvalidRangeAddressToStringTest()
        {
            var address = ProduceInvalidAddress();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToString(), Is.EqualTo("#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.A1), Is.EqualTo("#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.Default), Is.EqualTo("#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.R1C1), Is.EqualTo("'Sheet 1'!#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.Default, true), Is.EqualTo("'Sheet 1'!#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.R1C1, true), Is.EqualTo("'Sheet 1'!#REF!"));
            });
        }

        [Test]
        public void InvalidRangeAddressToStringFixedTest()
        {
            var address = ProduceInvalidAddress();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToStringFixed(), Is.EqualTo("#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1), Is.EqualTo("#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default), Is.EqualTo("#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1), Is.EqualTo("#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("'Sheet 1'!#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true), Is.EqualTo("'Sheet 1'!#REF!"));
            });
        }

        [Test]
        public void InvalidRangeAddressToStringRelativeTest()
        {
            var address = ProduceInvalidAddress();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToStringRelative(), Is.EqualTo("#REF!"));
                Assert.That(address.ToStringRelative(true), Is.EqualTo("'Sheet 1'!#REF!"));
            });
        }

        [Test]
        public void RangeAddressOnDeletedWorksheetToStringTest()
        {
            var address = ProduceAddressOnDeletedWorksheet();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToString(), Is.EqualTo("#REF!A1:B2"));
                Assert.That(address.ToString(XLReferenceStyle.A1), Is.EqualTo("#REF!A1:B2"));
                Assert.That(address.ToString(XLReferenceStyle.Default), Is.EqualTo("#REF!A1:B2"));
                Assert.That(address.ToString(XLReferenceStyle.R1C1), Is.EqualTo("#REF!R1C1:R2C2"));
                Assert.That(address.ToString(XLReferenceStyle.A1, true), Is.EqualTo("#REF!A1:B2"));
                Assert.That(address.ToString(XLReferenceStyle.Default, true), Is.EqualTo("#REF!A1:B2"));
                Assert.That(address.ToString(XLReferenceStyle.R1C1, true), Is.EqualTo("#REF!R1C1:R2C2"));
            });
        }

        [Test]
        public void RangeAddressOnDeletedWorksheetToStringFixedTest()
        {
            var address = ProduceAddressOnDeletedWorksheet();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToStringFixed(), Is.EqualTo("#REF!$A$1:$B$2"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1), Is.EqualTo("#REF!$A$1:$B$2"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default), Is.EqualTo("#REF!$A$1:$B$2"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1), Is.EqualTo("#REF!R1C1:R2C2"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true), Is.EqualTo("#REF!$A$1:$B$2"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("#REF!$A$1:$B$2"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true), Is.EqualTo("#REF!R1C1:R2C2"));
            });
        }

        [Test]
        public void RangeAddressOnDeletedWorksheetToStringRelativeTest()
        {
            var address = ProduceAddressOnDeletedWorksheet();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToStringRelative(), Is.EqualTo("#REF!A1:B2"));
                Assert.That(address.ToStringRelative(true), Is.EqualTo("#REF!A1:B2"));
            });
        }

        [Test]
        public void InvalidRangeAddressOnDeletedWorksheetToStringTest()
        {
            var address = ProduceInvalidAddressOnDeletedWorksheet();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToString(), Is.EqualTo("#REF!#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.A1), Is.EqualTo("#REF!#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.Default), Is.EqualTo("#REF!#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.R1C1), Is.EqualTo("#REF!#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.A1, true), Is.EqualTo("#REF!#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.Default, true), Is.EqualTo("#REF!#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.R1C1, true), Is.EqualTo("#REF!#REF!"));
            });
        }

        [Test]
        public void InvalidRangeAddressOnDeletedWorksheetToStringFixedTest()
        {
            var address = ProduceInvalidAddressOnDeletedWorksheet();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToStringFixed(), Is.EqualTo("#REF!#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1), Is.EqualTo("#REF!#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default), Is.EqualTo("#REF!#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1), Is.EqualTo("#REF!#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true), Is.EqualTo("#REF!#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("#REF!#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true), Is.EqualTo("#REF!#REF!"));
            });
        }

        [Test]
        public void InvalidRangeAddressOnDeletedWorksheetToStringRelativeTest()
        {
            var address = ProduceInvalidAddressOnDeletedWorksheet();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToStringRelative(), Is.EqualTo("#REF!#REF!"));
                Assert.That(address.ToStringRelative(true), Is.EqualTo("#REF!#REF!"));
            });
        }

        [Test]
        public void FullSpanAddressCannotChange()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            var wsRange = ws.AsRange();
            var row = ws.FirstRow().RowBelow(4).AsRange();
            var column = ws.FirstColumn().ColumnRight(4).AsRange();

            Assert.Multiple(() =>
            {
                Assert.That(wsRange.RangeAddress.ToString(), Is.EqualTo($"1:{XLHelper.MaxRowNumber}"));
                Assert.That(row.RangeAddress.ToString(), Is.EqualTo("5:5"));
                Assert.That(column.RangeAddress.ToString(), Is.EqualTo("E:E"));
            });

            ws.Columns("Y:Z").Delete();
            ws.Rows("9:10").Delete();

            Assert.Multiple(() =>
            {
                Assert.That(wsRange.RangeAddress.ToString(), Is.EqualTo($"1:{XLHelper.MaxRowNumber}"));
                Assert.That(row.RangeAddress.ToString(), Is.EqualTo("5:5"));
                Assert.That(column.RangeAddress.ToString(), Is.EqualTo("E:E"));
            });
        }

        [Test]
        public void RangeAddressIsNormalized()
        {
            var ws = new XLWorkbook().AddWorksheet();

            XLRangeAddress rangeAddress;

            rangeAddress = (XLRangeAddress)ws.Range(ws.Cell("A1"), ws.Cell("C3")).RangeAddress;
            Assert.That(rangeAddress.IsNormalized, Is.True);

            rangeAddress = (XLRangeAddress)ws.Range(ws.Cell("C3"), ws.Cell("A1")).RangeAddress;
            Assert.That(rangeAddress.IsNormalized, Is.False);

            rangeAddress = (XLRangeAddress)ws.Range("B2:B1").RangeAddress;
            Assert.That(rangeAddress.IsNormalized, Is.False);

            rangeAddress = (XLRangeAddress)ws.Range("B2:B10").RangeAddress;
            Assert.That(rangeAddress.IsNormalized, Is.True);

            rangeAddress = (XLRangeAddress)ws.Range("B:B").RangeAddress;
            Assert.That(rangeAddress.IsNormalized, Is.True);

            rangeAddress = (XLRangeAddress)ws.Range("2:2").RangeAddress;
            Assert.That(rangeAddress.IsNormalized, Is.True);

            rangeAddress = (XLRangeAddress)ws.RangeAddress;
            Assert.That(rangeAddress.IsNormalized, Is.True);
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

            Assert.Multiple(() =>
            {
                Assert.That(rangeAddress.IsValid, Is.True);
                Assert.That(rangeAddress.IsNormalized, Is.True);
            });
            Assert.Throws<InvalidOperationException>(() => rangeAddress.AsRange());

            var ws = new XLWorkbook().AddWorksheet() as XLWorksheet;
            rangeAddress = new XLRangeAddress
            (
                new XLAddress(ws, 1, 1, false, false),
                new XLAddress(ws, 5, 5, false, false)
            );

            Assert.Multiple(() =>
            {
                Assert.That(rangeAddress.IsValid, Is.True);
                Assert.That(rangeAddress.IsNormalized, Is.True);
            });
            Assert.DoesNotThrow(() => rangeAddress.AsRange());
        }

        [Test]
        public void RelativeRanges()
        {
            var ws = new XLWorkbook().AddWorksheet();

            IXLRangeAddress rangeAddress;

            rangeAddress = ws.Range("D4:E4").RangeAddress.Relative(ws.Range("A1:E4").RangeAddress, ws.Range("B10:F14").RangeAddress);
            Assert.Multiple(() =>
            {
                Assert.That(rangeAddress.IsValid, Is.True);
                Assert.That(rangeAddress.ToString(), Is.EqualTo("E13:F13"));
            });

            rangeAddress = ws.Range("D4:E4").RangeAddress.Relative(ws.Range("B10:F14").RangeAddress, ws.Range("A1:E4").RangeAddress);
            Assert.Multiple(() =>
            {
                Assert.That(rangeAddress.IsValid, Is.False);
                Assert.That(rangeAddress.ToString(), Is.EqualTo("#REF!"));
            });

            rangeAddress = ws.Range("C3").RangeAddress.Relative(ws.Range("A1:B2").RangeAddress, ws.Range("C3").RangeAddress);
            Assert.Multiple(() =>
            {
                Assert.That(rangeAddress.IsValid, Is.True);
                Assert.That(rangeAddress.ToString(), Is.EqualTo("E5:E5"));
            });

            rangeAddress = ws.Range("B2").RangeAddress.Relative(ws.Range("A1").RangeAddress, ws.Range("C3").RangeAddress);
            Assert.Multiple(() =>
            {
                Assert.That(rangeAddress.IsValid, Is.True);
                Assert.That(rangeAddress.ToString(), Is.EqualTo("D4:D4"));
            });

            rangeAddress = ws.Range("A1").RangeAddress.Relative(ws.Range("B2").RangeAddress, ws.Range("A1").RangeAddress);
            Assert.Multiple(() =>
            {
                Assert.That(rangeAddress.IsValid, Is.False);
                Assert.That(rangeAddress.ToString(), Is.EqualTo("#REF!"));
            });
        }

        [Test]
        public void TestSpanProperties()
        {
            var ws = new XLWorkbook().AddWorksheet() as XLWorksheet;

            var range = ws.Range("B3:E5");
            var rangeAddress = range.RangeAddress as IXLRangeAddress;
            Assert.Multiple(() =>
            {
                Assert.That(rangeAddress.ColumnSpan, Is.EqualTo(4));
                Assert.That(rangeAddress.RowSpan, Is.EqualTo(3));
                Assert.That(rangeAddress.NumberOfCells, Is.EqualTo(12));
            });

            range = ws.Range("E5:B3");
            rangeAddress = range.RangeAddress as IXLRangeAddress;
            Assert.Multiple(() =>
            {
                Assert.That(rangeAddress.ColumnSpan, Is.EqualTo(4));
                Assert.That(rangeAddress.RowSpan, Is.EqualTo(3));
                Assert.That(rangeAddress.NumberOfCells, Is.EqualTo(12));
            });

            rangeAddress = ProduceAddressOnDeletedWorksheet();
            Assert.Multiple(() =>
            {
                Assert.That(rangeAddress.ColumnSpan, Is.EqualTo(2));
                Assert.That(rangeAddress.RowSpan, Is.EqualTo(2));
                Assert.That(rangeAddress.NumberOfCells, Is.EqualTo(4));
            });

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
