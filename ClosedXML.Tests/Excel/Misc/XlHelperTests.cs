using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class XlHelperTests
    {
        private static void CheckColumnNumber(int column)
        {
            Assert.That(XLHelper.GetColumnNumberFromLetter(XLHelper.GetColumnLetterFromNumber(column)), Is.EqualTo(column));
        }

        [Test]
        public void InvalidA1Addresses()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLHelper.IsValidA1Address(""), Is.False);
                Assert.That(XLHelper.IsValidA1Address("A"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("a"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("1"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("-1"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("AAAA1"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("XFG1"), Is.False);

                Assert.That(XLHelper.IsValidA1Address("@A1"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("@AA1"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("@AAA1"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("[A1"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("[AA1"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("[AAA1"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("{A1"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("{AA1"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("{AAA1"), Is.False);

                Assert.That(XLHelper.IsValidA1Address("A1@"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("AA1@"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("AAA1@"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("A1["), Is.False);
                Assert.That(XLHelper.IsValidA1Address("AA1["), Is.False);
                Assert.That(XLHelper.IsValidA1Address("AAA1["), Is.False);
                Assert.That(XLHelper.IsValidA1Address("A1{"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("AA1{"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("AAA1{"), Is.False);

                Assert.That(XLHelper.IsValidA1Address("@A1@"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("@AA1@"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("@AAA1@"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("[A1["), Is.False);
                Assert.That(XLHelper.IsValidA1Address("[AA1["), Is.False);
                Assert.That(XLHelper.IsValidA1Address("[AAA1["), Is.False);
                Assert.That(XLHelper.IsValidA1Address("{A1{"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("{AA1{"), Is.False);
                Assert.That(XLHelper.IsValidA1Address("{AAA1{"), Is.False);
            });
        }

        [Test]
        public void PlusAA1_Is_Not_an_address()
        {
            Assert.That(XLHelper.IsValidA1Address("+AA1"), Is.False);
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
            CheckColumnNumber(1353);
        }

        [Test]
        public void ValidA1Addresses()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLHelper.IsValidA1Address("A1"), Is.True);
                Assert.That(XLHelper.IsValidA1Address("A" + XLHelper.MaxRowNumber), Is.True);
                Assert.That(XLHelper.IsValidA1Address("Z1"), Is.True);
                Assert.That(XLHelper.IsValidA1Address("Z" + XLHelper.MaxRowNumber), Is.True);

                Assert.That(XLHelper.IsValidA1Address("AA1"), Is.True);
                Assert.That(XLHelper.IsValidA1Address("AA" + XLHelper.MaxRowNumber), Is.True);
                Assert.That(XLHelper.IsValidA1Address("ZZ1"), Is.True);
                Assert.That(XLHelper.IsValidA1Address("ZZ" + XLHelper.MaxRowNumber), Is.True);

                Assert.That(XLHelper.IsValidA1Address("AAA1"), Is.True);
                Assert.That(XLHelper.IsValidA1Address("AAA" + XLHelper.MaxRowNumber), Is.True);
                Assert.That(XLHelper.IsValidA1Address(XLHelper.MaxColumnLetter + "1"), Is.True);
                Assert.That(XLHelper.IsValidA1Address(XLHelper.MaxColumnLetter + XLHelper.MaxRowNumber), Is.True);
            });
        }

        [Test]
        public void TestColumnLetterLookup()
        {
            var columnLetters = new List<string>();
            for (int c = 1; c <= XLHelper.MaxColumnNumber; c++)
            {
                var columnLetter = NaiveGetColumnLetterFromNumber(c);
                columnLetters.Add(columnLetter);

                Assert.That(XLHelper.GetColumnLetterFromNumber(c), Is.EqualTo(columnLetter));
            }

            foreach (var cl in columnLetters)
            {
                var columnNumber = NaiveGetColumnNumberFromLetter(cl);
                Assert.That(XLHelper.GetColumnNumberFromLetter(cl), Is.EqualTo(columnNumber));
            }
        }

        [TestCase("R")]
        [TestCase("C")]
        [TestCase("RC")]
        [TestCase("R111C222")]
        [TestCase("R[]C")]
        [TestCase("RC[]")]
        [TestCase("R[]C[]")]
        [TestCase("R[111]C222")]
        [TestCase("R111C[222]")]
        [TestCase("R[111]C[222]")]
        [TestCase("R[-111]C[-222]")]
        public void ValidRCAddresses(string address)
        {
            Assert.That(XLHelper.IsValidRCAddress(address), Is.True);
        }

        [TestCase("RD")]
        [TestCase("CC")]
        [TestCase("R[-]C222")]
        [TestCase("R[]C[-]")]
        [TestCase("_R111C222")]
        public void InvalidRCAddresses(string address)
        {
            Assert.That(XLHelper.IsValidRCAddress(address), Is.False);
        }

        #region Old XLHelper methods

        private static readonly string[] letters = new[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

        /// <summary>
        /// These used to be the methods in XLHelper, but were later changed
        /// We now use them as a check against the new methods
        /// Gets the column number of a given column letter.
        /// </summary>
        /// <param name="columnLetter"> The column letter to translate into a column number. </param>
        private static int NaiveGetColumnNumberFromLetter(string columnLetter)
        {
            if (string.IsNullOrEmpty(columnLetter)) throw new ArgumentNullException("columnLetter");

            int retVal;
            columnLetter = columnLetter.ToUpper();

            //Extra check because we allow users to pass row col positions in as strings
            if (columnLetter[0] <= '9')
            {
                retVal = int.Parse(columnLetter, XLHelper.NumberStyle, XLHelper.ParseCulture);
                return retVal;
            }

            int sum = 0;

            for (int i = 0; i < columnLetter.Length; i++)
            {
                sum *= 26;
                sum += (columnLetter[i] - 'A' + 1);
            }

            return sum;
        }

        /// <summary>
        /// Gets the column letter of a given column number.
        /// </summary>
        /// <param name="columnNumber">The column number to translate into a column letter.</param>
        /// <param name="trimToAllowed">if set to <c>true</c> the column letter will be restricted to the allowed range.</param>
        private static string NaiveGetColumnLetterFromNumber(int columnNumber, bool trimToAllowed = false)
        {
            if (trimToAllowed) columnNumber = XLHelper.TrimColumnNumber(columnNumber);

            columnNumber--; // Adjust for start on column 1
            if (columnNumber <= 25)
            {
                return letters[columnNumber];
            }
            var firstPart = (columnNumber) / 26;
            var remainder = ((columnNumber) % 26) + 1;
            return NaiveGetColumnLetterFromNumber(firstPart) + NaiveGetColumnLetterFromNumber(remainder);
        }

        #endregion Old XLHelper methods
    }
}
