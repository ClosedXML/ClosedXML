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

        [Theory]
        public void Abs_ReturnsItselfOnPositiveNumbers([Range(0, 10, 0.1)] double input)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ABS({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(input, actual, tolerance * 10);
        }

        [Theory]
        public void Abs_ReturnsTheCorrectValueOnNegativeInput([Range(-10, -0.1, 0.1)] double input)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ABS({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(-input, actual, tolerance * 10);
        }

        [TestCase(-10, 3.041924001)]
        [TestCase(-9, 3.030935432)]
        [TestCase(-8, 3.017237659)]
        [TestCase(-7, 2.999695599)]
        [TestCase(-6, 2.976443976)]
        [TestCase(-5, 2.944197094)]
        [TestCase(-4, 2.89661399)]
        [TestCase(-3, 2.819842099)]
        [TestCase(-2, 2.677945045)]
        [TestCase(-1, 2.35619449)]
        [TestCase(0, 1.570796327)]
        [TestCase(1, 0.785398163)]
        [TestCase(2, 0.463647609)]
        [TestCase(3, 0.321750554)]
        [TestCase(4, 0.244978663)]
        [TestCase(5, 0.19739556)]
        [TestCase(6, 0.165148677)]
        [TestCase(7, 0.141897055)]
        [TestCase(8, 0.124354995)]
        [TestCase(9, 0.110657221)]
        [TestCase(10, 0.099668652)]
        public void Acot_ReturnsCorrectValue(double input, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ACOT({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedResult, actual, tolerance * 10);
        }

        [Theory]
        public void Acos_ThrowsNumberExceptionOutsideRange([Range(1.1, 3, 0.1)] double input)
        {
            // checking input and it's additive inverse as both are outside range.
            Assert.Throws<NumberException>(() => XLWorkbook.EvaluateExpr(string.Format(@"ACOS({0})", input.ToString(CultureInfo.InvariantCulture))));
            Assert.Throws<NumberException>(() => XLWorkbook.EvaluateExpr(string.Format(@"ACOS({0})", (-input).ToString(CultureInfo.InvariantCulture))));
        }

        [TestCase(-1, 3.141592654)]
        [TestCase(-0.9, 2.690565842)]
        [TestCase(-0.8, 2.498091545)]
        [TestCase(-0.7, 2.346193823)]
        [TestCase(-0.6, 2.214297436)]
        [TestCase(-0.5, 2.094395102)]
        [TestCase(-0.4, 1.982313173)]
        [TestCase(-0.3, 1.875488981)]
        [TestCase(-0.2, 1.772154248)]
        [TestCase(-0.1, 1.670963748)]
        [TestCase(0, 1.570796327)]
        [TestCase(0.1, 1.470628906)]
        [TestCase(0.2, 1.369438406)]
        [TestCase(0.3, 1.266103673)]
        [TestCase(0.4, 1.159279481)]
        [TestCase(0.5, 1.047197551)]
        [TestCase(0.6, 0.927295218)]
        [TestCase(0.7, 0.79539883)]
        [TestCase(0.8, 0.643501109)]
        [TestCase(0.9, 0.451026812)]
        [TestCase(1, 0)]
        public void Acos_ReturnsCorrectValue(double input, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ACOS({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedResult, actual, tolerance * 10);
        }

        [Theory]
        public void Acosh_NumbersBelow1ThrowNumberException([Range(-1, 0.9, 0.1)] double input)
        {
            Assert.Throws<NumberException>(() => XLWorkbook.EvaluateExpr(string.Format(@"ACOSH({0})", input.ToString(CultureInfo.InvariantCulture))));
        }

        [TestCase(1.2, 0.622362504)]
        [TestCase(1.5, 0.96242365)]
        [TestCase(1.8, 1.192910731)]
        [TestCase(2.1, 1.372859144)]
        [TestCase(2.4, 1.522079367)]
        [TestCase(2.7, 1.650193455)]
        [TestCase(3, 1.762747174)]
        [TestCase(3.3, 1.863279351)]
        [TestCase(3.6, 1.954207529)]
        [TestCase(3.9, 2.037266466)]
        [TestCase(4.2, 2.113748231)]
        [TestCase(4.5, 2.184643792)]
        [TestCase(4.8, 2.250731414)]
        [TestCase(5.1, 2.312634419)]
        [TestCase(5.4, 2.370860342)]
        [TestCase(5.7, 2.425828318)]
        [TestCase(6, 2.47788873)]
        [TestCase(1, 0)]
        public void Acosh_ReturnsCorrectValue(double input, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ACOSH({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedResult, actual, tolerance * 10);
        }

        [TestCase(-10, -0.100335348)]
        [TestCase(-9, -0.111571776)]
        [TestCase(-8, -0.125657214)]
        [TestCase(-7, -0.143841036)]
        [TestCase(-6, -0.168236118)]
        [TestCase(-5, -0.202732554)]
        [TestCase(-4, -0.255412812)]
        [TestCase(-3, -0.34657359)]
        [TestCase(-2, -0.549306144)]
        [TestCase(2, 0.549306144)]
        [TestCase(3, 0.34657359)]
        [TestCase(4, 0.255412812)]
        [TestCase(5, 0.202732554)]
        [TestCase(6, 0.168236118)]
        [TestCase(7, 0.143841036)]
        [TestCase(8, 0.125657214)]
        [TestCase(9, 0.111571776)]
        [TestCase(10, 0.100335348)]
        public void Acoth_ReturnsCorrectValue(double input, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ACOTH({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedResult, actual, tolerance * 10);
        }

        [Theory]
        public void Acoth_ForPlusMinusXSmallerThan1_ThrowsNumberException([Range(-0.9, 0.9, 0.1)] double input)
        {
            Assert.Throws<NumberException>(() => XLWorkbook.EvaluateExpr(string.Format(@"ACOTH({0})", input.ToString(CultureInfo.InvariantCulture))));
        }

        [TestCase("LVII", 57)]
        [TestCase("mcmxii", 1912)]
        [TestCase("", 0)]
        [TestCase("-IV", -4)]
        [TestCase("   XIV", 14)]
        [TestCase("MCMLXXXIII ", 1983)]
        public void Arabic_ReturnsCorrectNumber(string roman, int arabic)
        {
            var actual = (int)XLWorkbook.EvaluateExpr(string.Format($"ARABIC(\"{roman}\")"));
            Assert.AreEqual(arabic, actual);
        }

        [Test]
        public void Arabic_ThrowsNumberExceptionOnMinus()
        {
            Assert.Throws<NumberException>(() => XLWorkbook.EvaluateExpr("ARABIC(\"-\")"));
        }

        [TestCase("- I")]
        [TestCase("roman")]
        public void Arabic_ThrowsValueExceptionOnInvalidNumber(string invalidRoman)
        {
            Assert.Throws<CellValueException>(() => XLWorkbook.EvaluateExpr($"ARABIC(\"{invalidRoman}\")"));
        }

        [Theory]
        public void Asin_ThrowsNumberExceptionWhenAbsOfInputGreaterThan1([Range(-3, -1.1, 0.1)] double input)
        {
            Assert.Throws<NumberException>(() => XLWorkbook.EvaluateExpr(string.Format(@"ASIN({0})", input.ToString(CultureInfo.InvariantCulture))));
            Assert.Throws<NumberException>(() => XLWorkbook.EvaluateExpr(string.Format(@"ASIN({0})", (-input).ToString(CultureInfo.InvariantCulture))));
        }

        [TestCase(-1, -1.570796327)]
        [TestCase(-0.9, -1.119769515)]
        [TestCase(-0.8, -0.927295218)]
        [TestCase(-0.7, -0.775397497)]
        [TestCase(-0.6, -0.643501109)]
        [TestCase(-0.5, -0.523598776)]
        [TestCase(-0.4, -0.411516846)]
        [TestCase(-0.3, -0.304692654)]
        [TestCase(-0.2, -0.201357921)]
        [TestCase(-0.1, -0.100167421)]
        [TestCase(0, 0)]
        [TestCase(0.1, 0.100167421)]
        [TestCase(0.2, 0.201357921)]
        [TestCase(0.3, 0.304692654)]
        [TestCase(0.4, 0.411516846)]
        [TestCase(0.5, 0.523598776)]
        [TestCase(0.6, 0.643501109)]
        [TestCase(0.7, 0.775397497)]
        [TestCase(0.8, 0.927295218)]
        [TestCase(0.9, 1.119769515)]
        [TestCase(1, 1.570796327)]
        public void Asin_ReturnsCorrectResult(double input, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ASIN({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedResult, actual, tolerance * 10);
        }

        [TestCase(0, 0)]
        [TestCase(0.1, 0.0998340788992076)]
        [TestCase(0.2, 0.198690110349241)]
        [TestCase(0.3, 0.295673047563422)]
        [TestCase(0.4, 0.390035319770715)]
        [TestCase(0.5, 0.481211825059603)]
        [TestCase(0.6, 0.568824898732248)]
        [TestCase(0.7, 0.652666566082356)]
        [TestCase(0.8, 0.732668256045411)]
        [TestCase(0.9, 0.808866935652783)]
        [TestCase(1, 0.881373587019543)]
        [TestCase(2, 1.44363547517881)]
        [TestCase(3, 1.81844645923207)]
        [TestCase(4, 2.0947125472611)]
        [TestCase(5, 2.31243834127275)]
        public void Asinh_ReturnsCorrectResult(double input, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ASINH({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedResult, actual, tolerance);
            var minusActual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ASINH({0})", (-input).ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(-expectedResult, minusActual, tolerance);
        }

        [TestCase(4, 3, 20)]
        [TestCase(10, 3, 220)]
        [TestCase(0, 0, 1)]
        public void Combina_CalculatesCorrectValues(int number, int chosen, int expectedResult)
        {
            var actualResult = XLWorkbook.EvaluateExpr($"COMBINA({number}, {chosen})");
            Assert.AreEqual(expectedResult, (long)actualResult);
        }

        [Theory]
        public void Combina_Returns1WhenChosenIs0([Range(0, 10)]int number)
        {
            Combina_CalculatesCorrectValues(number, 0, 1);
        }

        [TestCase(4.23, 3, 20)]
        [TestCase(10.4, 3.14, 220)]
        [TestCase(0, 0.4, 1)]
        public void Combina_TruncatesNumbersCorrectly(double number, double chosen, int expectedResult)
        {
            var actualResult = XLWorkbook.EvaluateExpr(string.Format(
                @"COMBINA({0}, {1})",
                number.ToString(CultureInfo.InvariantCulture),
                chosen.ToString(CultureInfo.InvariantCulture)));

            Assert.AreEqual(expectedResult, (long)actualResult);
        }

        [TestCase(-1, 2)]
        [TestCase(-3, -2)]
        [TestCase(2, -2)]
        public void Combina_ThrowsNumExceptionOnInvalidValues(int number, int chosen)
        {
            Assert.Throws<NumberException>(() => XLWorkbook.EvaluateExpr(
                string.Format(
                    @"COMBINA({0}, {1})",
                    number.ToString(CultureInfo.InvariantCulture),
                    chosen.ToString(CultureInfo.InvariantCulture))));
        }

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

        [TestCase(-10, -1.000000004)]
        [TestCase(-9, -1.00000003)]
        [TestCase(-8, -1.000000225)]
        [TestCase(-7, -1.000001663)]
        [TestCase(-6, -1.000012289)]
        [TestCase(-5, -1.000090804)]
        [TestCase(-4, -1.00067115)]
        [TestCase(-3, -1.004969823)]
        [TestCase(-2, -1.037314721)]
        [TestCase(-1, -1.313035285)]
        [TestCase(1, 1.313035285)]
        [TestCase(2, 1.037314721)]
        [TestCase(3, 1.004969823)]
        [TestCase(4, 1.00067115)]
        [TestCase(5, 1.000090804)]
        [TestCase(6, 1.000012289)]
        [TestCase(7, 1.000001663)]
        [TestCase(8, 1.000000225)]
        [TestCase(9, 1.00000003)]
        [TestCase(10, 1.000000004)]
        public void Coth_Examples(double input, double expected)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"COTH({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expected, actual, tolerance * 10.0);
        }

        [Test]
        public void Cot_On0_ThrowsDivisionByZeroException()
        {
            Assert.Throws<DivisionByZeroException>(() => XLWorkbook.EvaluateExpr(@"COTH(0)"));
        }

        [TestCase(-10, 1.838163961)]
        [TestCase(-9, -2.426486644)]
        [TestCase(-8, -1.010756218)]
        [TestCase(-7, -1.522101063)]
        [TestCase(-6, 3.578899547)]
        [TestCase(-5, 1.042835213)]
        [TestCase(-4, 1.321348709)]
        [TestCase(-3, -7.086167396)]
        [TestCase(-2, -1.09975017)]
        [TestCase(-1, -1.188395106)]
        [TestCase(1, 1.188395106)]
        [TestCase(2, 1.09975017)]
        [TestCase(3, 7.086167396)]
        [TestCase(4, -1.321348709)]
        [TestCase(5, -1.042835213)]
        [TestCase(6, -3.578899547)]
        [TestCase(7, 1.522101063)]
        [TestCase(8, 1.010756218)]
        [TestCase(9, 2.426486644)]
        [TestCase(10, -1.838163961)]
        public void Csc_ReturnsCorrectValues(double input, double expected)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"CSC({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expected, actual, tolerance * 10);
        }

        [Test]
        public void Csc_On0_ThrowsDivisionByZeroException()
        {
            Assert.Throws<DivisionByZeroException>(() => XLWorkbook.EvaluateExpr(@"CSC(0)"));
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

        [TestCase(0, 1)]
        [TestCase(0.3, 1.0467516)]
        [TestCase(0.6, 1.21162831)]
        [TestCase(0.9, 1.60872581)]
        [TestCase(1.2, 2.759703601)]
        [TestCase(1.5, 14.1368329)]
        [TestCase(1.8, -4.401367872)]
        [TestCase(2.1, -1.980801656)]
        [TestCase(2.4, -1.356127641)]
        [TestCase(2.7, -1.10610642)]
        [TestCase(3.0, -1.010108666)]
        [TestCase(3.3, -1.012678974)]
        [TestCase(3.6, -1.115127532)]
        [TestCase(3.9, -1.377538917)]
        [TestCase(4.2, -2.039730601)]
        [TestCase(4.5, -4.743927548)]
        [TestCase(4.8, 11.42870421)]
        [TestCase(5.1, 2.645658426)]
        [TestCase(5.4, 1.575565187)]
        [TestCase(5.7, 1.198016873)]
        [TestCase(6.0, 1.041481927)]
        [TestCase(6.3, 1.000141384)]
        [TestCase(6.6, 1.052373922)]
        [TestCase(6.9, 1.225903187)]
        [TestCase(7.2, 1.643787029)]
        [TestCase(7.5, 2.884876262)]
        [TestCase(7.8, 18.53381902)]
        [TestCase(8.1, -4.106031636)]
        [TestCase(8.4, -1.925711244)]
        [TestCase(8.7, -1.335743646)]
        [TestCase(9.0, -1.097537906)]
        [TestCase(9.3, -1.007835594)]
        [TestCase(9.6, -1.015550252)]
        [TestCase(9.9, -1.124617578)]
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

        [TestCase(-9, 0.00024682)]
        [TestCase(-8, 0.000670925)]
        [TestCase(-7, 0.001823762)]
        [TestCase(-6, 0.004957474)]
        [TestCase(-5, 0.013475282)]
        [TestCase(-4, 0.036618993)]
        [TestCase(-3, 0.099327927)]
        [TestCase(-2, 0.265802229)]
        [TestCase(-1, 0.648054274)]
        [TestCase(0, 1)]
        public void Sech_ReturnsCorrectNumber(double input, double expectedOutput)
        {
            double result = (double)XLWorkbook.EvaluateExpr(
                string.Format(
                    @"SECH({0})",
                    input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedOutput, result, 0.00001);

            // as the secant is symmetric for positive and negative numbers, let's assert twice:
            double resultForNegative = (double)XLWorkbook.EvaluateExpr(
                string.Format(
                    @"SECH({0})",
                    (-input).ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedOutput, resultForNegative, 0.00001);
        }

        /// <summary>
        /// refers to Example 1 from the Excel documentation,
        /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
        /// </summary>
        /// <param name="expectedOutcome"></param>
        /// <param name="formula"></param>
        [TestCase(63000, "SUMIF(A1:A4,\">160000\", B1:B4)")]
        [TestCase(900000, "SUMIF(A1:A4,\">160000\")")]
        [TestCase(21000, "SUMIF(A1:A4, 300000, B1:B4)")]
        [TestCase(28000, "SUMIF(A1:A4, \">\" &C1, B1:B4)")]
        public void SumIf_ReturnsCorrectValues_ReferenceExample1FromMicrosoft(int expectedOutcome, string formula)
        {
            using (var wb = new XLWorkbook())
            {
                wb.ReferenceStyle = XLReferenceStyle.A1;

                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell(1, 1).Value = 100000;
                ws.Cell(1, 2).Value = 7000;
                ws.Cell(2, 1).Value = 200000;
                ws.Cell(2, 2).Value = 14000;
                ws.Cell(3, 1).Value = 300000;
                ws.Cell(3, 2).Value = 21000;
                ws.Cell(4, 1).Value = 400000;
                ws.Cell(4, 2).Value = 28000;

                ws.Cell(1, 3).Value = 300000;

                Assert.AreEqual(expectedOutcome, (double)ws.Evaluate(formula));
            }
        }

        /// <summary>
        /// refers to Example 2 from the Excel documentation,
        /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
        /// </summary>
        /// <param name="expectedOutcome"></param>
        /// <param name="formula"></param>
        [TestCase(2000, "SUMIF(A2:A7,\"Fruits\", C2:C7)")]
        [TestCase(12000, "SUMIF(A2:A7,\"Vegetables\", C2:C7)")]
        [TestCase(4300, "SUMIF(B2:B7, \"*es\", C2:C7)")]
        [TestCase(400, "SUMIF(A2:A7, \"\", C2:C7)")]
        public void SumIf_ReturnsCorrectValues_ReferenceExample2FromMicrosoft(int expectedOutcome, string formula)
        {
            using (var wb = new XLWorkbook())
            {
                wb.ReferenceStyle = XLReferenceStyle.A1;

                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell(2, 1).Value = "Vegetables";
                ws.Cell(3, 1).Value = "Vegetables";
                ws.Cell(4, 1).Value = "Fruits";
                ws.Cell(5, 1).Value = "";
                ws.Cell(6, 1).Value = "Vegetables";
                ws.Cell(7, 1).Value = "Fruits";

                ws.Cell(2, 2).Value = "Tomatoes";
                ws.Cell(3, 2).Value = "Celery";
                ws.Cell(4, 2).Value = "Oranges";
                ws.Cell(5, 2).Value = "Butter";
                ws.Cell(6, 2).Value = "Carrots";
                ws.Cell(7, 2).Value = "Apples";

                ws.Cell(2, 3).Value = 2300;
                ws.Cell(3, 3).Value = 5500;
                ws.Cell(4, 3).Value = 800;
                ws.Cell(5, 3).Value = 400;
                ws.Cell(6, 3).Value = 4200;
                ws.Cell(7, 3).Value = 1200;

                ws.Cell(1, 3).Value = 300000;

                Assert.AreEqual(expectedOutcome, (double)ws.Evaluate(formula));
            }
        }

        /// <summary>
        /// refers to Example 1 to SumIf from the Excel documentation.
        /// As SumIfs should behave the same if called with three parameters, we can take that example here again.
        /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
        /// </summary>
        /// <param name="expectedOutcome"></param>
        /// <param name="formula"></param>
        [TestCase(63000, "SUMIFS(B1:B4, \">160000\", A1:A4)")]
        [TestCase(21000, "SUMIFS(B1:B4, 300000, A1:A4)")]
        [TestCase(28000, "SUMIFS(B1:B4, \">\" &C1, A1:A4)")]
        public void SumIfs_ReturnsCorrectValues_ReferenceExampleForSumIf1FromMicrosoft(int expectedOutcome, string formula)
        {
            using (var wb = new XLWorkbook())
            {
                wb.ReferenceStyle = XLReferenceStyle.A1;

                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell(1, 1).Value = 100000;
                ws.Cell(1, 2).Value = 7000;
                ws.Cell(2, 1).Value = 200000;
                ws.Cell(2, 2).Value = 14000;
                ws.Cell(3, 1).Value = 300000;
                ws.Cell(3, 2).Value = 21000;
                ws.Cell(4, 1).Value = 400000;
                ws.Cell(4, 2).Value = 28000;

                ws.Cell(1, 3).Value = 300000;

                Assert.AreEqual(expectedOutcome, (double)ws.Evaluate(formula));
            }
        }

        /// <summary>
        /// refers to Example 2 to SumIf from the Excel documentation.
        /// As SumIfs should behave the same if called with three parameters, we can take that example here again.
        /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
        /// </summary>
        /// <param name="expectedOutcome"></param>
        /// <param name="formula"></param>
        [TestCase(2000, "SUMIFS(C2:C7, \"Fruits\", A2:A7)")]
        [TestCase(12000, "SUMIFS(C2:C7, \"Vegetables\", A2:A7)")]
        [TestCase(4300, "SUMIFS(C2:C7, \"*es\", B2:B7)")]
        [TestCase(400, "SUMIFS(C2:C7, \"\", A2:A7)")]
        public void SumIfs_ReturnsCorrectValues_ReferenceExample2FromMicrosoft(int expectedOutcome, string formula)
        {
            using (var wb = new XLWorkbook())
            {
                wb.ReferenceStyle = XLReferenceStyle.A1;

                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell(2, 1).Value = "Vegetables";
                ws.Cell(3, 1).Value = "Vegetables";
                ws.Cell(4, 1).Value = "Fruits";
                ws.Cell(5, 1).Value = "";
                ws.Cell(6, 1).Value = "Vegetables";
                ws.Cell(7, 1).Value = "Fruits";

                ws.Cell(2, 2).Value = "Tomatoes";
                ws.Cell(3, 2).Value = "Celery";
                ws.Cell(4, 2).Value = "Oranges";
                ws.Cell(5, 2).Value = "Butter";
                ws.Cell(6, 2).Value = "Carrots";
                ws.Cell(7, 2).Value = "Apples";

                ws.Cell(2, 3).Value = 2300;
                ws.Cell(3, 3).Value = 5500;
                ws.Cell(4, 3).Value = 800;
                ws.Cell(5, 3).Value = 400;
                ws.Cell(6, 3).Value = 4200;
                ws.Cell(7, 3).Value = 1200;

                ws.Cell(1, 3).Value = 300000;

                Assert.AreEqual(expectedOutcome, (double)ws.Evaluate(formula));
            }
        }

        /// <summary>
        /// refers to example data and formula to SumIfs in the Excel documentation,
        /// <see cref="https://support.office.com/en-us/article/SUMIFS-function-c9e748f5-7ea7-455d-9406-611cebce642b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
        /// </summary>
        [TestCase(20, "=SUMIFS(A2:A9, B2:B9, \"=A*\", C2:C9, \"Tom\")")]
        [TestCase(30, "=SUMIFS(A2:A9, B2:B9, \"<>Bananas\", C2:C9, \"Tom\")")]
        public void SumIfs_ReturnsCorrectValues_ReferenceExampleFromMicrosoft(
            int result,
            string formula)
        {
            using (var wb = new XLWorkbook())
            {
                wb.ReferenceStyle = XLReferenceStyle.A1;
                var ws = wb.AddWorksheet("Sheet1");

                ws.Cell(1, 1).Value = 5;
                ws.Cell(1, 2).Value = "Apples";
                ws.Cell(1, 3).Value = "Tom";

                ws.Cell(2, 1).Value = 4;
                ws.Cell(2, 2).Value = "Apples";
                ws.Cell(2, 3).Value = "Sarah";

                ws.Cell(3, 1).Value = 15;
                ws.Cell(3, 2).Value = "Artichokes";
                ws.Cell(3, 3).Value = "Tom";

                ws.Cell(4, 1).Value = 3;
                ws.Cell(4, 2).Value = "Artichokes";
                ws.Cell(4, 3).Value = "Sarah";

                ws.Cell(5, 1).Value = 22;
                ws.Cell(5, 2).Value = "Bananas";
                ws.Cell(5, 3).Value = "Tom";

                ws.Cell(6, 1).Value = 12;
                ws.Cell(6, 2).Value = "Bananas";
                ws.Cell(6, 3).Value = "Sarah";

                ws.Cell(7, 1).Value = 10;
                ws.Cell(7, 2).Value = "Carrots";
                ws.Cell(7, 3).Value = "Tom";

                ws.Cell(8, 1).Value = 33;
                ws.Cell(8, 2).Value = "Carrots";
                ws.Cell(8, 3).Value = "Sarah";
            }
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
