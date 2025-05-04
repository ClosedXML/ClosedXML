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

            Assert.Multiple(() =>
            {
                Assert.That(address.ToString(), Is.EqualTo("A1"));
                Assert.That(address.ToString(XLReferenceStyle.A1), Is.EqualTo("A1"));
                Assert.That(address.ToString(XLReferenceStyle.R1C1), Is.EqualTo("R1C1"));
                Assert.That(address.ToString(XLReferenceStyle.Default), Is.EqualTo("A1"));
                Assert.That(address.ToString(XLReferenceStyle.Default, true), Is.EqualTo("Sheet1!A1"));

                Assert.That(address.ToStringRelative(), Is.EqualTo("A1"));
                Assert.That(address.ToStringRelative(true), Is.EqualTo("Sheet1!A1"));

                Assert.That(address.ToStringFixed(), Is.EqualTo("$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1), Is.EqualTo("$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1), Is.EqualTo("R1C1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default), Is.EqualTo("$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true), Is.EqualTo("Sheet1!$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true), Is.EqualTo("Sheet1!R1C1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("Sheet1!$A$1"));
            });
        }

        [Test]
        public void ToStringTestWithSpace()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet 1");
            IXLAddress address = ws.Cell(1, 1).Address;

            Assert.Multiple(() =>
            {
                Assert.That(address.ToString(), Is.EqualTo("A1"));
                Assert.That(address.ToString(XLReferenceStyle.A1), Is.EqualTo("A1"));
                Assert.That(address.ToString(XLReferenceStyle.R1C1), Is.EqualTo("R1C1"));
                Assert.That(address.ToString(XLReferenceStyle.Default), Is.EqualTo("A1"));
                Assert.That(address.ToString(XLReferenceStyle.Default, true), Is.EqualTo("'Sheet 1'!A1"));

                Assert.That(address.ToStringRelative(), Is.EqualTo("A1"));
                Assert.That(address.ToStringRelative(true), Is.EqualTo("'Sheet 1'!A1"));

                Assert.That(address.ToStringFixed(), Is.EqualTo("$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1), Is.EqualTo("$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1), Is.EqualTo("R1C1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default), Is.EqualTo("$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true), Is.EqualTo("'Sheet 1'!R1C1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("'Sheet 1'!$A$1"));
            });
        }

        [Test]
        public void InvalidAddressToStringTest()
        {
            var address = ProduceInvalidAddress();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToString(), Is.EqualTo("#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.A1), Is.EqualTo("#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.R1C1), Is.EqualTo("#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.Default), Is.EqualTo("#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.Default, true), Is.EqualTo("'Sheet 1'!#REF!"));
            });
        }

        [Test]
        public void InvalidAddressToStringFixedTest()
        {
            var address = ProduceInvalidAddress();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToStringFixed(), Is.EqualTo("#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1), Is.EqualTo("#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1), Is.EqualTo("#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default), Is.EqualTo("#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true), Is.EqualTo("'Sheet 1'!#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("'Sheet 1'!#REF!"));
            });
        }

        [Test]
        public void InvalidAddressToStringRelativeTest()
        {
            var address = ProduceInvalidAddress();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToStringRelative(), Is.EqualTo("#REF!"));
                Assert.That(address.ToStringRelative(true), Is.EqualTo("'Sheet 1'!#REF!"));
            });
        }

        [Test]
        public void AddressOnDeletedWorksheetToStringTest()
        {
            var address = ProduceAddressOnDeletedWorksheet();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToString(), Is.EqualTo("A1"));
                Assert.That(address.ToString(XLReferenceStyle.A1), Is.EqualTo("A1"));
                Assert.That(address.ToString(XLReferenceStyle.R1C1), Is.EqualTo("R1C1"));
                Assert.That(address.ToString(XLReferenceStyle.Default), Is.EqualTo("A1"));
                Assert.That(address.ToString(XLReferenceStyle.Default, true), Is.EqualTo("#REF!A1"));
            });
        }

        [Test]
        public void AddressOnDeletedWorksheetToStringFixedTest()
        {
            var address = ProduceAddressOnDeletedWorksheet();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToStringFixed(), Is.EqualTo("$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1), Is.EqualTo("$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1), Is.EqualTo("R1C1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default), Is.EqualTo("$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true), Is.EqualTo("#REF!$A$1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true), Is.EqualTo("#REF!R1C1"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("#REF!$A$1"));
            });
        }

        [Test]
        public void AddressOnDeletedWorksheetToStringRelativeTest()
        {
            var address = ProduceAddressOnDeletedWorksheet();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToStringRelative(), Is.EqualTo("A1"));
                Assert.That(address.ToStringRelative(true), Is.EqualTo("#REF!A1"));
            });
        }

        [Test]
        public void InvalidAddressOnDeletedWorksheetToStringTest()
        {
            var address = ProduceInvalidAddressOnDeletedWorksheet();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToString(), Is.EqualTo("#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.A1), Is.EqualTo("#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.R1C1), Is.EqualTo("#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.Default), Is.EqualTo("#REF!"));
                Assert.That(address.ToString(XLReferenceStyle.Default, true), Is.EqualTo("#REF!#REF!"));
            });
        }

        [Test]
        public void InvalidAddressOnDeletedWorksheetToStringFixedTest()
        {
            var address = ProduceInvalidAddressOnDeletedWorksheet();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToStringFixed(), Is.EqualTo("#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1), Is.EqualTo("#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1), Is.EqualTo("#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default), Is.EqualTo("#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true), Is.EqualTo("#REF!#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true), Is.EqualTo("#REF!#REF!"));
                Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("#REF!#REF!"));
            });
        }

        [Test]
        public void InvalidAddressOnDeletedWorksheetToStringRelativeTest()
        {
            var address = ProduceInvalidAddressOnDeletedWorksheet();

            Assert.Multiple(() =>
            {
                Assert.That(address.ToStringRelative(), Is.EqualTo("#REF!"));
                Assert.That(address.ToStringRelative(true), Is.EqualTo("#REF!#REF!"));
            });
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
