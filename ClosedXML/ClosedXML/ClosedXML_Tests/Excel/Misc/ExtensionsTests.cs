using System;
using ClosedXML.Excel;
using NUnit.Framework;

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
    }
}