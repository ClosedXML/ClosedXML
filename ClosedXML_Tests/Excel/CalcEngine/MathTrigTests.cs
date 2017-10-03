using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using ClosedXML.Excel.CalcEngine.Exceptions;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML_Tests.Excel.CalcEngine
{
    [TestFixture]
    public class MathTrigTests
    {
        private readonly double tolerance = 1e-10;

        [TestCase("FF", 16, 255)]
        [TestCase("111", 2, 7)]
        [TestCase("zap", 36, 45745)]
        public void Decimal(string inputString, int radix, int expectedResult)
        {
            var actualResult = XLWorkbook.EvaluateExpr($"DECIMAL(\"{inputString}\", {radix})");
            Assert.AreEqual(expectedResult, actualResult);
        }

        [Test]
        public void Decimal_ZeroIsZeroInAnyRadix([Range(1, 36)] int radix)
        {
            Assert.AreEqual(0, XLWorkbook.EvaluateExpr($"DECIMAL(\"0\", {radix})"));
        }

        [Theory]
        public void Decimal_ReturnsErrorForRadiansGreater36([Range(37, 255)] int radix)
        {
            Assert.Throws<NumberException>(() => XLWorkbook.EvaluateExpr($"DECIMAL(\"0\", {radix})"));
        }

        [Theory]
        public void Decimal_ReturnsErrorForRadiansSmaller2([Range(-5, 1)] int radix)
        {
            Assert.Throws<NumberException>(() => XLWorkbook.EvaluateExpr($"DECIMAL(\"0\", {radix})"));
        }

        [Test]
        public void Floor()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"FLOOR(1.2)");
            Assert.AreEqual(1, actual);

            actual = XLWorkbook.EvaluateExpr(@"FLOOR(1.7)");
            Assert.AreEqual(1, actual);

            actual = XLWorkbook.EvaluateExpr(@"FLOOR(-1.7)");
            Assert.AreEqual(-2, actual);

            actual = XLWorkbook.EvaluateExpr(@"FLOOR(1.2, 1)");
            Assert.AreEqual(1, actual);

            actual = XLWorkbook.EvaluateExpr(@"FLOOR(1.7, 1)");
            Assert.AreEqual(1, actual);

            actual = XLWorkbook.EvaluateExpr(@"FLOOR(-1.7, 1)");
            Assert.AreEqual(-2, actual);

            actual = XLWorkbook.EvaluateExpr(@"FLOOR(0.4, 2)");
            Assert.AreEqual(0, actual);

            actual = XLWorkbook.EvaluateExpr(@"FLOOR(2.7, 2)");
            Assert.AreEqual(2, actual);

            actual = XLWorkbook.EvaluateExpr(@"FLOOR(7.8, 2)");
            Assert.AreEqual(6, actual);

            actual = XLWorkbook.EvaluateExpr(@"FLOOR(-5.5, -2)");
            Assert.AreEqual(-4, actual);
        }

        [Test]
        // Functions have to support a period first before we can implement this
        public void FloorMath()
        {
            double actual;

            actual = (double)XLWorkbook.EvaluateExpr(@"FLOOR.MATH(24.3, 5)");
            Assert.AreEqual(20, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"FLOOR.MATH(6.7)");
            Assert.AreEqual(6, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"FLOOR.MATH(-8.1, 2)");
            Assert.AreEqual(-10, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"FLOOR.MATH(5.5, 2.1, 0)");
            Assert.AreEqual(4.2, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"FLOOR.MATH(5.5, -2.1, 0)");
            Assert.AreEqual(4.2, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"FLOOR.MATH(5.5, 2.1, -1)");
            Assert.AreEqual(4.2, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"FLOOR.MATH(5.5, -2.1, -1)");
            Assert.AreEqual(4.2, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"FLOOR.MATH(-5.5, 2.1, 0)");
            Assert.AreEqual(-6.3, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"FLOOR.MATH(-5.5, -2.1, 0)");
            Assert.AreEqual(-6.3, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"FLOOR.MATH(-5.5, 2.1, -1)");
            Assert.AreEqual(-4.2, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"FLOOR.MATH(-5.5, -2.1, -1)");
            Assert.AreEqual(-4.2, actual, tolerance);
        }

        [Test]
        public void Mod()
        {
            double actual;

            actual = (double)XLWorkbook.EvaluateExpr(@"MOD(1.5, 1)");
            Assert.AreEqual(0.5, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"MOD(3, 2)");
            Assert.AreEqual(1, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"MOD(-3, 2)");
            Assert.AreEqual(1, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"MOD(3, -2)");
            Assert.AreEqual(-1, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"MOD(-3, -2)");
            Assert.AreEqual(-1, actual, tolerance);

            //////

            actual = (double)XLWorkbook.EvaluateExpr(@"MOD(-4.3, -0.5)");
            Assert.AreEqual(-0.3, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"MOD(6.9, -0.2)");
            Assert.AreEqual(-0.1, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"MOD(0.7, 0.6)");
            Assert.AreEqual(0.1, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr(@"MOD(6.2, 1.1)");
            Assert.AreEqual(0.7, actual, tolerance);
        }

        [Test]
        public void SumProduct()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                ws.FirstCell().Value = Enumerable.Range(1, 10);
                ws.FirstCell().CellRight().Value = Enumerable.Range(1, 10).Reverse();

                Assert.AreEqual(2, ws.Evaluate("SUMPRODUCT(A2)"));
                Assert.AreEqual(55, ws.Evaluate("SUMPRODUCT(A1:A10)"));
                Assert.AreEqual(220, ws.Evaluate("SUMPRODUCT(A1:A10, B1:B10)"));

                Assert.Throws<NoValueAvailableException>(() => ws.Evaluate("SUMPRODUCT(A1:A10, B1:B5)"));
            }
        }
    }
}
