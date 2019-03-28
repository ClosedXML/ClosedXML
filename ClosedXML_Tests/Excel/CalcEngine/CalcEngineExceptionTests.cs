using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using ClosedXML.Excel.CalcEngine.Exceptions;
using NUnit.Framework;
using System;
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
            Assert.Throws<DivisionByZeroException>(() => new XLWorkbook().AddWorksheet().Evaluate("0/0"));
        }

        [Test]
        public void InvalidFunction()
        {
            Exception ex;
            ex = Assert.Throws<ExpressionParseException>(() => XLWorkbook.EvaluateExpr("XXX(A1:A2)"));
            Assert.That(ex.Message, Is.EqualTo("Unknown function: XXX"));

            var ws = new XLWorkbook().AddWorksheet();
            ex = Assert.Throws<ExpressionParseException>(() => ws.Evaluate("XXX(A1:A2)"));
            Assert.That(ex.Message, Is.EqualTo("Unknown function: XXX"));
        }
    }
}
