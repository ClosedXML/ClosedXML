using System;
using System.Collections.Generic;
using System.Text;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel.Formula
{
    public class XLReferenceTests
    {
        #region XLCellReference

        [TestCase(-5, 6, true, false)]
        [TestCase(0, 6, true, false)]
        [TestCase(5, -6, false, true)]
        [TestCase(5, 0, false, true)]
        [TestCase(XLHelper.MaxRowNumber + 1, 1, false, false)]
        [TestCase(XLHelper.MaxRowNumber + 1, 1, true, false)]
        [TestCase(-XLHelper.MaxRowNumber - 1, 1, false, false)]
        [TestCase(1, XLHelper.MaxColumnNumber + 1, false, false)]
        [TestCase(1, XLHelper.MaxColumnNumber + 1, false, true)]
        [TestCase(1, -XLHelper.MaxColumnNumber - 1, false, false)]
        public void XLCellReference_Invalid(int row, int column, bool rowIsAbsolute, bool columnIsAbsolute)
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => new XLCellReference(row, column, rowIsAbsolute, columnIsAbsolute));
        }

        [TestCase(-5, -6, false, false, "G14", "A9")]
        [TestCase(-5, 0, false, false, "G14", "G9")]
        [TestCase(0, 0, false, false, "G14", "G14")]
        [TestCase(5, 6, false, false, "G14", "M19")]
        [TestCase(5, 6, true, false, "G14", "M$5")]
        [TestCase(5, 6, true, true, "G14", "$F$5")]
        [TestCase(-5, -6, false, false, "F14", "#REF!")]
        [TestCase(-5, -6, false, false, "G4", "#REF!")]
        public void XLCellReference_ToStringA1(int row, int column, bool rowIsAbsolute, bool columnIsAbsolute,
            string baseAddressA1, string expectedA1)
        {
            var cellReference = new XLCellReference(row, column, rowIsAbsolute, columnIsAbsolute);
            var baseAddress = XLAddress.Create(baseAddressA1);
            var actualA1 = cellReference.ToStringA1(baseAddress);

            Assert.AreEqual(expectedA1, actualA1);
        }

        [TestCase(-5, -6, false, false, "R[-5]C[-6]")]
        [TestCase(-5, 0, false, false, "R[-5]C")]
        [TestCase(0, 0, false, false, "RC")]
        [TestCase(5, 6, false, false, "R[5]C[6]")]
        [TestCase(5, 6, true, false, "R5C[6]")]
        [TestCase(5, 6, true, true, "R5C6")]
        public void XLCellReference_ToStringR1C1_Valid(int row, int column, bool rowIsAbsolute, bool columnIsAbsolute, string expectedR1C1)
        {
            var cellReference = new XLCellReference(row, column, rowIsAbsolute, columnIsAbsolute);

            var actualR1C1 = cellReference.ToStringR1C1();

            Assert.AreEqual(expectedR1C1, actualR1C1);
        }
        #endregion XLCellReference

        #region XLRangeReference

        [Test]
        public void XLRangeReference_ToStringA1()
        {
            var cellReference1 = new XLCellReference(0, 0, false, false);
            var cellReference2 = new XLCellReference(5, 6, false, false);
            var rangeReference = new XLRangeReference(cellReference1, cellReference2);
            var baseAddress = XLAddress.Create("G14");

            var actualA1 = rangeReference.ToStringA1(baseAddress);
            Assert.AreEqual("G14:M19", actualA1);
        }

        [Test]
        public void XLRangeReference_ToStringR1C1()
        {
            var cellReference1 = new XLCellReference(0, 0, false, false);
            var cellReference2 = new XLCellReference(5, 6, false, false);
            var rangeReference = new XLRangeReference(cellReference1, cellReference2);

            var actualR1C1 = rangeReference.ToStringR1C1();

            Assert.AreEqual("RC:R[5]C[6]", actualR1C1);
        }

        [Test]
        public void XLRangeReferenceThrows_WhenPassedNull()
        {
            var cellReference = new XLCellReference(0, 0, false, false);

            Assert.Throws<ArgumentNullException>(() => new XLRangeReference(null, cellReference));
            Assert.Throws<ArgumentNullException>(() => new XLRangeReference(cellReference, null));
            Assert.DoesNotThrow(() => new XLRangeReference(cellReference, cellReference));
        }
        #endregion XLRangeReference

        #region XLRowReference

        [TestCase(-5, true)]
        [TestCase(0, true)]
        [TestCase(XLHelper.MaxRowNumber + 1, false)]
        [TestCase(XLHelper.MaxRowNumber + 1, true)]
        [TestCase(-XLHelper.MaxRowNumber - 1, false)]
        public void XLRowReference_Invalid(int row, bool rowIsAbsolute)
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => new XLRowReference(row, rowIsAbsolute));
        }

        [TestCase(-5, false, "G14", "9:9")]
        [TestCase(0, false, "G14", "14:14")]
        [TestCase(5, false, "G14", "19:19")]
        [TestCase(5, true, "G14", "$5:$5")]
        [TestCase(-5, false, "G4", "#REF!")]
        public void XLRowReference_ToStringA1(int row, bool rowIsAbsolute, string baseAddressA1, string expectedA1)
        {
            var rowReference = new XLRowReference(row, rowIsAbsolute);
            var baseAddress = XLAddress.Create(baseAddressA1);
            var actualA1 = rowReference.ToStringA1(baseAddress);

            Assert.AreEqual(expectedA1, actualA1);
        }

        [TestCase(-5, false, "R[-5]")]
        [TestCase(0, false, "R")]
        [TestCase(5, false, "R[5]")]
        [TestCase(5, true, "R5")]
        public void XLRowReference_ToStringR1C1_Valid(int row, bool rowIsAbsolute, string expectedR1C1)
        {
            var rowReference = new XLRowReference(row, rowIsAbsolute);

            var actualR1C1 = rowReference.ToStringR1C1();

            Assert.AreEqual(expectedR1C1, actualR1C1);
        }
        #endregion XLRowReference

        #region XLRowRangeReference

        [TestCase(0, 0, false, false, "G14", "14:14")]
        [TestCase(-1, 1, false, false, "G14", "13:15")]
        [TestCase(1, 10, true, false, "G14", "$1:24")]
        [TestCase(1, 10, true, true, "G14", "$1:$10")]
        [TestCase(-5, 6, false, true, "G4", "#REF!")]
        public void XLRowRangeReference_ToStringA1(int firstRow, int lastRow, bool firstRowIsAbsolute, bool lastRowIsAbsolute,
            string baseAddressA1, string expectedA1)
        {
            var firstRowReference = new XLRowReference(firstRow, firstRowIsAbsolute);
            var lastRowReference = new XLRowReference(lastRow, lastRowIsAbsolute);
            var rowRangeReference = new XLRowRangeReference(firstRowReference, lastRowReference);
            var baseAddress = XLAddress.Create(baseAddressA1);

            var actualA1 = rowRangeReference.ToStringA1(baseAddress);
            Assert.AreEqual(expectedA1, actualA1);
        }

        [TestCase(0, 0, false, false, "R:R")]
        [TestCase(-1, 1, false, false, "R[-1]:R[1]")]
        [TestCase(1, 10, true, false, "R1:R[10]")]
        [TestCase(1, 10, true, true, "R1:R10")]
        public void XLRowRangeReference_ToStringR1C1(int firstRow, int lastRow, bool firstRowIsAbsolute, bool lastRowIsAbsolute, string expectedR1C1)
        {
            var firstRowReference = new XLRowReference(firstRow, firstRowIsAbsolute);
            var lastRowReference = new XLRowReference(lastRow, lastRowIsAbsolute);
            var rowRangeReference = new XLRowRangeReference(firstRowReference, lastRowReference);

            var actualR1C1 = rowRangeReference.ToStringR1C1();

            Assert.AreEqual(expectedR1C1, actualR1C1);
        }

        [Test]
        public void XLRowRangeReferenceThrows_WhenPassedNull()
        {
            var rowReference = new XLRowReference(0, rowIsAbsolute: false);

            Assert.Throws<ArgumentNullException>(() => new XLRowRangeReference(null, rowReference));
            Assert.Throws<ArgumentNullException>(() => new XLRowRangeReference(rowReference, null));
            Assert.DoesNotThrow(() => new XLRowRangeReference(rowReference, rowReference));
        }
        #endregion XLRowRangeReference

        #region XLColumnReference

        [TestCase(-5, true)]
        [TestCase(0, true)]
        [TestCase(XLHelper.MaxColumnNumber + 1, false)]
        [TestCase(XLHelper.MaxColumnNumber + 1, true)]
        [TestCase(-XLHelper.MaxColumnNumber - 1, false)]
        public void XLColumnReference_Invalid(int column, bool columnIsAbsolute)
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => new XLColumnReference(column, columnIsAbsolute));
        }

        [TestCase(-6, false, "G14", "A:A")]
        [TestCase(0, false, "G14", "G:G")]
        [TestCase(6, false, "G14", "M:M")]
        [TestCase(6, true, "G14", "$F:$F")]
        [TestCase(-6, false, "F14", "#REF!")]
        public void XLColumnReference_ToStringA1(int column, bool columnIsAbsolute, string baseAddressA1, string expectedA1)
        {
            var columnReference = new XLColumnReference(column, columnIsAbsolute);
            var baseAddress = XLAddress.Create(baseAddressA1);
            var actualA1 = columnReference.ToStringA1(baseAddress);

            Assert.AreEqual(expectedA1, actualA1);
        }

        [TestCase(-6, false, "C[-6]")]
        [TestCase(0, false, "C")]
        [TestCase(6, false, "C[6]")]
        [TestCase(6, true, "C6")]
        public void XLColumnReference_ToStringR1C1_Valid(int column, bool columnIsAbsolute, string expectedR1C1)
        {
            var columnReference = new XLColumnReference(column, columnIsAbsolute);

            var actualR1C1 = columnReference.ToStringR1C1();

            Assert.AreEqual(expectedR1C1, actualR1C1);
        }
        #endregion XLRowReference

    }
}
