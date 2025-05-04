using ClosedXML.Excel;
using NUnit.Framework;
using System.Globalization;
using System.Threading;

namespace ClosedXML.Tests.Excel.CalcEngine
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
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr("CHAR(-2)"), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(XLWorkbook.EvaluateExpr("CHAR(270)"), Is.EqualTo(XLError.IncompatibleValue));
            });
        }

        [Test]
        public void DivisionByZero()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr("0/0"), Is.EqualTo(XLError.DivisionByZero));
                Assert.That(new XLWorkbook().AddWorksheet().Evaluate("0/0"), Is.EqualTo(XLError.DivisionByZero));
            });
        }

        [Test]
        public void InvalidFunction()
        {
            Assert.That(XLWorkbook.EvaluateExpr("XXX(A1:A2)"), Is.EqualTo(XLError.NameNotRecognized));

            var ws = new XLWorkbook().AddWorksheet();
            Assert.That(ws.Evaluate("XXX(A1:A2)"), Is.EqualTo(XLError.NameNotRecognized));
        }

        [Test]
        public void NestedNameNotRecognizedException()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").SetFormulaA1("=XXX");
            ws.Cell("A2").SetFormulaA1(@"=IFERROR(A1, ""Success"")");

            Assert.That(ws.Cell("A2").Value, Is.EqualTo("Success"));
        }
    }
}
