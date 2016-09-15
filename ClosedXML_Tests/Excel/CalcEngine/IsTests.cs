using System;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel.CalcEngine
{

    [TestFixture]
    public class IsTests
    {
        [Test]
        public void IsBlank_true()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            var actual = ws.Evaluate("=IsBlank(A1)");
            Assert.AreEqual(true, actual);
        }

        [Test]
        public void IsBlank_false()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            ws.Cell("A1").Value = " ";
            var actual = ws.Evaluate("=IsBlank(A1)");
            Assert.AreEqual(false, actual);
        }
    }
}