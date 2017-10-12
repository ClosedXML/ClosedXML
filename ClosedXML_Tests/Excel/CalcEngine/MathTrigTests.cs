using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine.Exceptions;
using NUnit.Framework;
using System;
using System.Globalization;
using System.Linq;

namespace ClosedXML_Tests.Excel.CalcEngine
{
    [TestFixture]
    public class MathTrigTests
    {
        private readonly double tolerance = 1e-10;

        [TestCase(1, 0.642092616)]
        [TestCase(2, -0.457657554)]
        [TestCase(3, -7.015252551)]
        [TestCase(4, 0.863691154)]
        [TestCase(5, -0.295812916)]
        [TestCase(6, -3.436353004)]
        [TestCase(7, 1.147515422)]
        [TestCase(8, -0.147065064)]
        [TestCase(9, -2.210845411)]
        [TestCase(10, 1.542351045)]
        [TestCase(11, -0.004425741)]
        [TestCase(Math.PI * 0.5, 0)]
        [TestCase(45, 0.617369624)]
        [TestCase(-2, 0.457657554)]
        [TestCase(-3, 7.015252551)]
        public void Cot(double input, double expected)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"COT({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expected, actual, tolerance * 10.0);
        }

        [Test]
        public void Cot_Input0()
        {
            Assert.Throws<DivisionByZeroException>(() => XLWorkbook.EvaluateExpr("COT(0)"));
        }

        [TestCase("FF", 16, 255)]
        [TestCase("111", 2, 7)]
        [TestCase("zap", 36, 45745)]
        public void Decimal(string inputString, int radix, int expectedResult)
        {
            var actualResult = XLWorkbook.EvaluateExpr($"DECIMAL(\"{inputString}\", {radix})");
            Assert.AreEqual(expectedResult, actualResult);
        }

        [Test]
        public void Decimal_ZeroIsZeroInAnyRadix([Range(2, 36)] int radix)
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

        [Theory]
        public void Multinomial_AnySingleValue_ReturnsOne([Range(0, 10, 0.1)] double number)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(
                @"MULTINOMIAL({0})",
                number.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(1, actual);
        }

        [TestCase(1, 2, 3)]
        [TestCase(2, 3, 10)]
        [TestCase(3, 4, 35)]
        [TestCase(4, 5, 126)]
        [TestCase(5, 6, 462)]
        [TestCase(6, 7, 1716)]
        [TestCase(7, 8, 6435)]
        [TestCase(8, 9, 24310)]
        [TestCase(9, 10, 92378)]
        [TestCase(10, 11, 352716)]
        [TestCase(11, 12, 1352078)]
        [TestCase(12, 13, 5200300)]
        [TestCase(13, 14, 20058300)]
        [TestCase(14, 15, 77558760)]
        [TestCase(15, 16, 300540195)]
        public void Multinomial_TwoValuesReturnCorrectValue(long first, long second, long expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(
                @"MULTINOMIAL({0}, {1})",
                first.ToString(CultureInfo.InvariantCulture),
                second.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedResult, actual, Math.Pow(10, -5));
        }

        [TestCase(1, 1)]
        [TestCase(2, 3)]
        [TestCase(3, 60)]
        [TestCase(4, 12600)]
        [TestCase(5, 37837800)]
        [TestCase(6, 2053230379200)]
        public void Multinomial_Values1ToNReturnCorrectValue(int last, long expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(
                @"MULTINOMIAL({0})",
                string.Join(
                    ", ",
                    Enumerable.Range(1, last)
                        .Select(i => i.ToString(CultureInfo.InvariantCulture)))));

            Assert.AreEqual(expectedResult, actual, 0.01);
        }

        [Theory]
        public void Multinomial_AnyNegativeValueThrowsNumberException([Range(-10, -1)] int x)
        {
            Assert.Throws<NumberException>(() => XLWorkbook.EvaluateExpr(
                string.Format(
                    @"MULTINOMIAL({0})",
                    string.Join(
                        ", ",
                        Enumerable.Range(x, 2)
                            .Reverse()
                            .Select(i => i.ToString(CultureInfo.InvariantCulture))))));
        }

        [Test]
        public void Multinomial_NonNumericValueThrowsNameNotRecognizedException()
        {
            Assert.Throws<NameNotRecognizedException>(
                () => XLWorkbook.EvaluateExpr(@"MULTINOMIAL(x)"));
        }

        [TestCase(   0, 1)]
        [TestCase( 0.3, 1.0467516)]
        [TestCase( 0.6, 1.21162831)]
        [TestCase( 0.9, 1.60872581)]
        [TestCase( 1.2, 2.759703601)]
        [TestCase( 1.5, 14.1368329)]
        [TestCase( 1.8, -4.401367872)]
        [TestCase( 2.1, -1.980801656)]
        [TestCase( 2.4, -1.356127641)]
        [TestCase( 2.7, -1.10610642)]
        [TestCase( 3.0, -1.010108666)]
        [TestCase( 3.3, -1.012678974)]
        [TestCase( 3.6, -1.115127532)]
        [TestCase( 3.9, -1.377538917)]
        [TestCase( 4.2, -2.039730601)]
        [TestCase( 4.5, -4.743927548)]
        [TestCase( 4.8, 11.42870421)]
        [TestCase( 5.1, 2.645658426)]
        [TestCase( 5.4, 1.575565187)]
        [TestCase( 5.7, 1.198016873)]
        [TestCase( 6.0, 1.041481927)]
        [TestCase( 6.3, 1.000141384)]
        [TestCase( 6.6, 1.052373922)]
        [TestCase( 6.9, 1.225903187)]
        [TestCase( 7.2, 1.643787029)]
        [TestCase( 7.5, 2.884876262)]
        [TestCase( 7.8, 18.53381902)]
        [TestCase( 8.1, -4.106031636)]
        [TestCase( 8.4, -1.925711244)]
        [TestCase( 8.7, -1.335743646)]
        [TestCase( 9.0, -1.097537906)]
        [TestCase( 9.3, -1.007835594)]
        [TestCase( 9.6, -1.015550252)]
        [TestCase( 9.9, -1.124617578)]
        [TestCase(10.2, -1.400039323)]
        [TestCase(10.5, -2.102886109)]
        [TestCase(10.8, -5.145888341)]
        [TestCase(11.1, 9.593612018)]
        [TestCase(11.4, 2.541355049)]
        [TestCase(45, 1.90359)]
        [TestCase(30, 6.48292)]
        public void Sec_ReturnsCorrectNumber(double input, double expectedOutput)
        {
            double result = (double)XLWorkbook.EvaluateExpr(
                string.Format(
                    @"SEC({0})",
                    input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedOutput, result, 0.00001);

            // as the secant is symmetric for positive and negative numbers, let's assert twice:
            double resultForNegative = (double)XLWorkbook.EvaluateExpr(
                string.Format(
                    @"SEC({0})",
                    (-input).ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedOutput, resultForNegative, 0.00001);
        }

        [Test]
        public void Sec_ThrowsCellValueExceptionOnNonNumericValue()
        {
            Assert.Throws<CellValueException>(() => XLWorkbook.EvaluateExpr(
                string.Format(
                    @"SEC(number)")));
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

        [TestCase(1, 0.850918128)]
        [TestCase(2, 0.275720565)]
        [TestCase(3, 0.09982157)]
        [TestCase(4, 0.03664357)]
        [TestCase(5, 0.013476506)]
        [TestCase(6, 0.004957535)]
        [TestCase(7, 0.001823765)]
        [TestCase(8, 0.000670925)]
        [TestCase(9, 0.00024682)]
        [TestCase(10, 0.000090799859712122200000)]
        [TestCase(11, 0.0000334034)]
        public void CSch_CalculatesCorrectValues(double input, double expectedOutput)
        {
            Assert.AreEqual(expectedOutput, (double)XLWorkbook.EvaluateExpr($@"CSCH({input})"), 0.000000001);
        }

        [Test]
        public void Csch_ReturnsDivisionByZeroErrorOnInput0()
        {
            Assert.Throws<DivisionByZeroException>(() => XLWorkbook.EvaluateExpr("CSCH(0)"));
        }

        [TestCase(8.9, 8)]
        [TestCase(-8.9, -9)]
        public void Int(double input, double expected)
        {
            var actual = XLWorkbook.EvaluateExpr(string.Format(@"INT({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expected, actual);

        }
    }
}
