using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
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
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("CHAR(-2)"));
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("CHAR(270)"));
        }

        [Test]
        public void DivisionByZero()
        {
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("0/0"));
            Assert.AreEqual(XLError.DivisionByZero, new XLWorkbook().AddWorksheet().Evaluate("0/0"));
        }

        [Test]
        public void InvalidFunction()
        {
            Assert.AreEqual(XLError.NameNotRecognized, XLWorkbook.EvaluateExpr("XXX(A1:A2)"));

            var ws = new XLWorkbook().AddWorksheet();
            Assert.AreEqual(XLError.NameNotRecognized, ws.Evaluate("XXX(A1:A2)"));
        }

        [Test]
        public void NestedNameNotRecognizedException()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").SetFormulaA1("=XXX");
            ws.Cell("A2").SetFormulaA1(@"=IFERROR(A1, ""Success"")");

            Assert.AreEqual("Success", ws.Cell("A2").Value);
        }
    }
}
