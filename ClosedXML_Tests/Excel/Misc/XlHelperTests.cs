using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel
{
    /// <summary>
    ///     Summary description for UnitTest1
    /// </summary>
    [TestFixture]
    public class XlHelperTests
    {
        private static void CheckColumnNumber(int column)
        {
            Assert.AreEqual(column, XLHelper.GetColumnNumberFromLetter(XLHelper.GetColumnLetterFromNumber(column)));
        }

        [Test]
        public void InvalidA1Addresses()
        {
            Assert.IsFalse(XLHelper.IsValidA1Address(""));
            Assert.IsFalse(XLHelper.IsValidA1Address("A"));
            Assert.IsFalse(XLHelper.IsValidA1Address("a"));
            Assert.IsFalse(XLHelper.IsValidA1Address("1"));
            Assert.IsFalse(XLHelper.IsValidA1Address("-1"));
            Assert.IsFalse(XLHelper.IsValidA1Address("AAAA1"));
            Assert.IsFalse(XLHelper.IsValidA1Address("XFG1"));

            Assert.IsFalse(XLHelper.IsValidA1Address("@A1"));
            Assert.IsFalse(XLHelper.IsValidA1Address("@AA1"));
            Assert.IsFalse(XLHelper.IsValidA1Address("@AAA1"));
            Assert.IsFalse(XLHelper.IsValidA1Address("[A1"));
            Assert.IsFalse(XLHelper.IsValidA1Address("[AA1"));
            Assert.IsFalse(XLHelper.IsValidA1Address("[AAA1"));
            Assert.IsFalse(XLHelper.IsValidA1Address("{A1"));
            Assert.IsFalse(XLHelper.IsValidA1Address("{AA1"));
            Assert.IsFalse(XLHelper.IsValidA1Address("{AAA1"));

            Assert.IsFalse(XLHelper.IsValidA1Address("A1@"));
            Assert.IsFalse(XLHelper.IsValidA1Address("AA1@"));
            Assert.IsFalse(XLHelper.IsValidA1Address("AAA1@"));
            Assert.IsFalse(XLHelper.IsValidA1Address("A1["));
            Assert.IsFalse(XLHelper.IsValidA1Address("AA1["));
            Assert.IsFalse(XLHelper.IsValidA1Address("AAA1["));
            Assert.IsFalse(XLHelper.IsValidA1Address("A1{"));
            Assert.IsFalse(XLHelper.IsValidA1Address("AA1{"));
            Assert.IsFalse(XLHelper.IsValidA1Address("AAA1{"));

            Assert.IsFalse(XLHelper.IsValidA1Address("@A1@"));
            Assert.IsFalse(XLHelper.IsValidA1Address("@AA1@"));
            Assert.IsFalse(XLHelper.IsValidA1Address("@AAA1@"));
            Assert.IsFalse(XLHelper.IsValidA1Address("[A1["));
            Assert.IsFalse(XLHelper.IsValidA1Address("[AA1["));
            Assert.IsFalse(XLHelper.IsValidA1Address("[AAA1["));
            Assert.IsFalse(XLHelper.IsValidA1Address("{A1{"));
            Assert.IsFalse(XLHelper.IsValidA1Address("{AA1{"));
            Assert.IsFalse(XLHelper.IsValidA1Address("{AAA1{"));
        }

        [Test]
        public void PlusAA1_Is_Not_an_address()
        {
            Assert.IsFalse(XLHelper.IsValidA1Address("+AA1"));
        }

        [Test]
        public void TestConvertColumnLetterToNumberAnd()
        {
            CheckColumnNumber(1);
            CheckColumnNumber(27);
            CheckColumnNumber(28);
            CheckColumnNumber(52);
            CheckColumnNumber(53);
            CheckColumnNumber(1000);
        }

        [Test]
        public void ValidA1Addresses()
        {
            Assert.IsTrue(XLHelper.IsValidA1Address("A1"));
            Assert.IsTrue(XLHelper.IsValidA1Address("A" + XLHelper.MaxRowNumber));
            Assert.IsTrue(XLHelper.IsValidA1Address("Z1"));
            Assert.IsTrue(XLHelper.IsValidA1Address("Z" + XLHelper.MaxRowNumber));

            Assert.IsTrue(XLHelper.IsValidA1Address("AA1"));
            Assert.IsTrue(XLHelper.IsValidA1Address("AA" + XLHelper.MaxRowNumber));
            Assert.IsTrue(XLHelper.IsValidA1Address("ZZ1"));
            Assert.IsTrue(XLHelper.IsValidA1Address("ZZ" + XLHelper.MaxRowNumber));

            Assert.IsTrue(XLHelper.IsValidA1Address("AAA1"));
            Assert.IsTrue(XLHelper.IsValidA1Address("AAA" + XLHelper.MaxRowNumber));
            Assert.IsTrue(XLHelper.IsValidA1Address(XLHelper.MaxColumnLetter + "1"));
            Assert.IsTrue(XLHelper.IsValidA1Address(XLHelper.MaxColumnLetter + XLHelper.MaxRowNumber));
        }
    }
}