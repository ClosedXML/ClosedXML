using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using NUnit.Framework;
using System;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class ExtensionsTests
    {
        [Test]
        public void FixNewLines()
        {
            Assert.That(Environment.NewLine, Is.EqualTo("\n".FixNewLines()));
            Assert.Multiple(() =>
            {
                Assert.That(Environment.NewLine, Is.EqualTo("\r\n".FixNewLines()));
                Assert.That("\rS" + Environment.NewLine, Is.EqualTo("\rS\n".FixNewLines()));
            });
            Assert.That(Environment.NewLine + Environment.NewLine, Is.EqualTo("\r\n\n".FixNewLines()));
        }

        [Test]
        public void DoubleSaveRound()
        {
            Double value = 1234.1234567;
            Assert.That(Math.Round(value, 6), Is.EqualTo(value.SaveRound()));
        }

        [Test]
        public void DoubleValueSaveRound()
        {
            Double value = 1234.1234567;
            Assert.That(Math.Round(value, 6), Is.EqualTo(new DoubleValue(value).SaveRound().Value));
        }

        [TestCase("NoEscaping", ExpectedResult = "NoEscaping")]
        [TestCase("1", ExpectedResult = "'1'")]
        [TestCase("AB-CD", ExpectedResult = "'AB-CD'")]
        [TestCase(" AB", ExpectedResult = "' AB'")]
        [TestCase("Test sheet", ExpectedResult = "'Test sheet'")]
        [TestCase("O'Kelly", ExpectedResult = "'O''Kelly'")]
        [TestCase("A2+3", ExpectedResult = "'A2+3'")]
        [TestCase("A\"B", ExpectedResult = "'A\"B'")]
        [TestCase("A!B", ExpectedResult = "'A!B'")]
        [TestCase("A~B", ExpectedResult = "'A~B'")]
        [TestCase("A^B", ExpectedResult = "'A^B'")]
        [TestCase("A&B", ExpectedResult = "'A&B'")]
        [TestCase("A>B", ExpectedResult = "'A>B'")]
        [TestCase("A<B", ExpectedResult = "'A<B'")]
        [TestCase("A.B", ExpectedResult = "A.B")]
        [TestCase(".", ExpectedResult = "'.'")]
        [TestCase("A_B", ExpectedResult = "A_B")]
        [TestCase("_", ExpectedResult = "_")]
        [TestCase("=", ExpectedResult = "'='")]
        [TestCase("A,B", ExpectedResult = "'A,B'")]
        [TestCase("A@B", ExpectedResult = "'A@B'")]
        [TestCase("(Test)", ExpectedResult = "'(Test)'")]
        [TestCase("A#", ExpectedResult = "'A#'")]
        [TestCase("A$", ExpectedResult = "'A$'")]
        [TestCase("A%", ExpectedResult = "'A%'")]
        [TestCase("ABC1", ExpectedResult = "'ABC1'")]
        [TestCase("ABCD1", ExpectedResult = "ABCD1")]
        [TestCase("R1C1", ExpectedResult = "'R1C1'")]
        [TestCase("A{", ExpectedResult = "'A{'")]
        [TestCase("A}", ExpectedResult = "'A}'")]
        [TestCase("A`", ExpectedResult = "'A`'")]
        [TestCase("Русский", ExpectedResult = "Русский")]
        [TestCase("日本語", ExpectedResult = "日本語")]
        [TestCase("한국어", ExpectedResult = "한국어")]
        [TestCase("Slovenščina", ExpectedResult = "Slovenščina")]
        [TestCase("", ExpectedResult = "")]
        [TestCase(null, ExpectedResult = null)]
        public string CanEscapeSheetName(string sheetName)
        {
            return StringExtensions.EscapeSheetName(sheetName);
        }

        [TestCase("TestSheet", ExpectedResult = "TestSheet")]
        [TestCase("'Test sheet'", ExpectedResult = "Test sheet")]
        [TestCase("'O''Kelly'", ExpectedResult = "O'Kelly")]
        public string CanUnescapeSheetName(string sheetName)
        {
            return StringExtensions.UnescapeSheetName(sheetName);
        }
    }
}
