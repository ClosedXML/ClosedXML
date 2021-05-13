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
    public class ExpressionTests
    {
        [OneTimeSetUp]
        public void SetCultureInfo()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
        }

        [Test]
        public void UnaryExpression()
        {
            // String
            Assert.AreEqual(XLWorkbook.EvaluateExpr("=+\"TestString\""), "TestString"); // Plus
            Assert.Throws<CellValueException>(() => XLWorkbook.EvaluateExpr("=-\"TestString\"")); // Minus

            // Double
            Assert.AreEqual(XLWorkbook.EvaluateExpr("+2.1"), 2.1); // Plus
            Assert.AreEqual(XLWorkbook.EvaluateExpr("-2.1"), -2.1); // Minus

            // Boolean True
            Assert.AreEqual(XLWorkbook.EvaluateExpr("+TRUE"), true); // Plus
            Assert.AreEqual(XLWorkbook.EvaluateExpr("-TRUE"), -1); // Minus

            // Boolean False
            Assert.AreEqual(XLWorkbook.EvaluateExpr("+FALSE"), false); // Plus
            Assert.AreEqual(XLWorkbook.EvaluateExpr("-FALSE"), 0); // Minus

            // DateTime
            Assert.AreEqual(XLWorkbook.EvaluateExpr("+TODAY()"), DateTime.Today); // Plus
            Assert.AreEqual(XLWorkbook.EvaluateExpr("-TODAY()"), -DateTime.Today.ToOADate()); // Minus
        }
    }
}
