using System;
using ClosedXML.Excel;
using NUnit.Framework;
using DocumentFormat.OpenXml;

namespace ClosedXML_Tests.Excel
{
    /// <summary>
    ///     Summary description for UnitTest1
    /// </summary>
    [TestFixture]
    public class ExtensionsTests
    {
        [Test]
        public void FixNewLines()
        {
            Assert.AreEqual("\n".FixNewLines(), Environment.NewLine);
            Assert.AreEqual("\r\n".FixNewLines(), Environment.NewLine);
            Assert.AreEqual("\rS\n".FixNewLines(), "\rS" + Environment.NewLine);
            Assert.AreEqual("\r\n\n".FixNewLines(), Environment.NewLine + Environment.NewLine);
        }

        [Test]
        public void DoubleSaveRound()
        {
            Double value = 1234.1234567;
            Assert.AreEqual(value.SaveRound(), Math.Round(value, 6));
        }

        [Test]
        public void DoubleValueSaveRound()
        {
            Double value = 1234.1234567;
            Assert.AreEqual(new DoubleValue(value).SaveRound().Value, Math.Round(value, 6));
        }
    }
}