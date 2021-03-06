using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine.Exceptions;
using NUnit.Framework;
using System;
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
            Assert.AreEqual(XLCalculationErrorType.CellValue, XLWorkbook.EvaluateExpr("CHAR(-2)"));
            Assert.AreEqual(XLCalculationErrorType.CellValue, XLWorkbook.EvaluateExpr("CHAR(270)"));
        }

        [Test]
        public void DivisionByZero()
        {
            Assert.AreEqual(XLCalculationErrorType.DivisionByZero, XLWorkbook.EvaluateExpr("0/0"));
            Assert.AreEqual(XLCalculationErrorType.DivisionByZero, new XLWorkbook().AddWorksheet().Evaluate("0/0"));
        }

        [Test]
        public void InvalidFunction()
        {
            Exception ex;
            ex = Assert.Throws<NameNotRecognizedException>(() => XLWorkbook.EvaluateExpr("XXX(A1:A2)"));
            Assert.That(ex.Message, Is.EqualTo("The identifier `XXX` was not recognised."));

            var ws = new XLWorkbook().AddWorksheet();
            ex = Assert.Throws<NameNotRecognizedException>(() => ws.Evaluate("XXX(A1:A2)"));
            Assert.That(ex.Message, Is.EqualTo("The identifier `XXX` was not recognised."));
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
