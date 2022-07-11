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
            Assert.Throws<CellValueException>(() => XLWorkbook.EvaluateExpr("CHAR(-2)"));
            Assert.Throws<CellValueException>(() => XLWorkbook.EvaluateExpr("CHAR(270)"));
        }

        [Test]
        public void DivisionByZero()
        {
            Assert.Throws<DivisionByZeroException>(() => XLWorkbook.EvaluateExpr("0/0"));
            Assert.Throws<DivisionByZeroException>(() => new XLWorkbook().AddWorksheet().Evaluate("0/0"));
        }

        [Test]
        public void InvalidFunction()
        {
            Exception ex;
            ex = Assert.Throws<NotImplementedException>(() => XLWorkbook.EvaluateExpr("XXX(A1:A2)"));
            Assert.That(ex.Message, Is.EqualTo("Evaluation of custom functions is not implemented."));

            var ws = new XLWorkbook().AddWorksheet();
            ex = Assert.Throws<NotImplementedException>(() => ws.Evaluate("XXX(A1:A2)"));
            Assert.That(ex.Message, Is.EqualTo("Evaluation of custom functions is not implemented."));
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
