using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine.Exceptions;
using NUnit.Framework;
using System.Globalization;
using System.Threading;

namespace ClosedXML_Tests.Excel.CalcEngine
{
    [TestFixture]
    public class CalcEngineExceptionTests
    {
        [OneTimeSetUp]
        public void SetCultureInfo()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
        }

        [Test]
        public void InvalidCharNumber()
        {
            Assert.Throws<CellValueException>(() => XLWorkbook.EvaluateExpr("CHAR(-2)"));
            Assert.Throws<CellValueException>(() => XLWorkbook.EvaluateExpr("CHAR(270)"));
        }

        [Test]
        public void DivisionByZero()
        {
            Assert.Throws<DivisionByZeroException>(() => XLWorkbook.EvaluateExpr("0/0"));
        }
    }
}
