// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine.Exceptions;
using NUnit.Framework;
using System;
using System.Globalization;
using System.Linq;

namespace ClosedXML.Tests.Excel.CalcEngine
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
        public void Acos_ThrowsNumberExceptionOutsideRange([Range(1.1, 3, 0.1)] double input)
        {
            // checking input and it's additive inverse as both are outside range.
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(string.Format(@"ACOS({0})", input.ToString(CultureInfo.InvariantCulture))));
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(string.Format(@"ACOS({0})", (-input).ToString(CultureInfo.InvariantCulture))));
        }

        [Theory]
        public void Acosh_NumbersBelow1ThrowNumberException([Range(-1, 0.9, 0.1)] double input)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(string.Format(@"ACOSH({0})", input.ToString(CultureInfo.InvariantCulture))));
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
        public void Acoth_ForPlusMinusXSmallerThan1_ThrowsNumberException([Range(-0.9, 0.9, 0.1)] double input)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(string.Format(@"ACOTH({0})", input.ToString(CultureInfo.InvariantCulture))));
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
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr("ARABIC(\"-\")"));
        }

        [TestCase("- I")]
        [TestCase("roman")]
        public void Arabic_ThrowsValueExceptionOnInvalidNumber(string invalidRoman)
        {
            Assert.AreEqual(XLCalculationErrorType.CellValue, XLWorkbook.EvaluateExpr($"ARABIC(\"{invalidRoman}\")"));
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

        [Theory]
        public void Asin_ThrowsNumberExceptionWhenAbsOfInputGreaterThan1([Range(-3, -1.1, 0.1)] double input)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(string.Format(@"ASIN({0})", input.ToString(CultureInfo.InvariantCulture))));
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(string.Format(@"ASIN({0})", (-input).ToString(CultureInfo.InvariantCulture))));
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

        [TestCase(0, 0)]
        [TestCase(0.1, 0.099668652491162)]
        [TestCase(0.2, 0.197395559849881)]
        [TestCase(0.3, 0.291456794477867)]
        [TestCase(0.4, 0.380506377112365)]
        [TestCase(0.5, 0.463647609000806)]
        [TestCase(0.6, 0.540419500270584)]
        [TestCase(0.7, 0.610725964389209)]
        [TestCase(0.8, 0.674740942223553)]
        [TestCase(0.9, 0.732815101786507)]
        [TestCase(1, 0.785398163397448)]
        [TestCase(2, 1.10714871779409)]
        [TestCase(3, 1.24904577239825)]
        [TestCase(4, 1.32581766366803)]
        [TestCase(5, 1.37340076694502)]
        public void Atan_ReturnsCorrectResult(double input, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ATAN({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedResult, actual, tolerance);
            var minusActual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ATAN({0})", (-input).ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(-expectedResult, minusActual, tolerance);
        }

        [Test]
        public void Atan2_Returns0OnSecond0AndFirstGreater0([Range(0.1, 5, 0.4)] double input)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ATAN2({0}, 0)", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(0, actual, tolerance);
        }

        [TestCase(1, 2, 1.10714871779409)]
        [TestCase(1, 3, 1.24904577239825)]
        [TestCase(2, 3, 0.98279372324733)]
        [TestCase(1, 4, 1.32581766366803)]
        [TestCase(3, 4, 0.92729521800161)]
        [TestCase(1, 5, 1.37340076694502)]
        [TestCase(2, 5, 1.19028994968253)]
        [TestCase(3, 5, 1.03037682652431)]
        [TestCase(4, 5, 0.89605538457134)]
        [TestCase(1, 6, 1.40564764938027)]
        [TestCase(5, 6, 0.87605805059819)]
        [TestCase(1, 7, 1.42889927219073)]
        [TestCase(2, 7, 1.29249666778979)]
        [TestCase(3, 7, 1.16590454050981)]
        [TestCase(4, 7, 1.05165021254837)]
        [TestCase(5, 7, 0.95054684081208)]
        [TestCase(6, 7, 0.86217005466723)]
        public void Atan2_ReturnsCorrectResults_EqualOnAllMultiplesOfFraction(double x, double y, double expectedResult)
        {
            for (int i = 1; i < 5; i++)
            {
                var actual = (double)XLWorkbook.EvaluateExpr(
                string.Format(
                    @"ATAN2({0}, {1})",
                    (x * i).ToString(CultureInfo.InvariantCulture),
                    (y * i).ToString(CultureInfo.InvariantCulture)));

                Assert.AreEqual(expectedResult, actual, tolerance);
            }
        }

        [Test]
        public void Atan2_ReturnsHalfPiOn0AsFirstInputWhenSecondGreater0([Range(0.1, 5, 0.4)] double input)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ATAN2(0, {0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(0.5 * Math.PI, actual, tolerance);
        }

        [Test]
        public void Atan2_ReturnsMinus3QuartersOfPiWhenFirstSmaller0AndSecondItsNegative([Range(-5, -0.1, 0.3)] double input)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ATAN2({0}, {0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(-0.75 * Math.PI, actual, tolerance);
        }

        [Test]
        public void Atan2_ReturnsMinusHalfPiOn0AsFirstInputWhenSecondSmaller0([Range(-5, -0.1, 0.4)] double input)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ATAN2(0, {0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(-0.5 * Math.PI, actual, tolerance);
        }

        [Test]
        public void Atan2_ReturnsPiOn0AsSecondInputWhenFirstSmaller0([Range(-5, -0.1, 0.4)] double input)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ATAN2({0}, 0)", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(Math.PI, actual, tolerance);
        }

        [Test]
        public void Atan2_ReturnsQuarterOfPiWhenInputsAreEqualAndGreater0([Range(0.1, 5, 0.3)] double input)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ATAN2({0}, {0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(0.25 * Math.PI, actual, tolerance);
        }

        [Test]
        public void Atan2_ThrowsDiv0ExceptionOn0And0()
        {
            Assert.AreEqual(XLCalculationErrorType.DivisionByZero, XLWorkbook.EvaluateExpr(@"ATAN2(0, 0)"));
        }

        [TestCase(-0.99, -2.64665241236225)]
        [TestCase(-0.9, -1.47221948958322)]
        [TestCase(-0.8, -1.09861228866811)]
        [TestCase(-0.6, -0.693147180559945)]
        [TestCase(-0.4, -0.423648930193602)]
        [TestCase(-0.2, -0.202732554054082)]
        [TestCase(0, 0)]
        [TestCase(0.2, 0.202732554054082)]
        [TestCase(0.4, 0.423648930193602)]
        [TestCase(0.6, 0.693147180559945)]
        [TestCase(0.8, 1.09861228866811)]
        [TestCase(-0.9, -1.47221948958322)]
        [TestCase(-0.990, -2.64665241236225)]
        [TestCase(-0.999, -3.8002011672502)]
        public void Atanh_ReturnsCorrectResults(double input, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(
                string.Format(
                    @"ATANH({0})",
                    input.ToString(CultureInfo.InvariantCulture)));

            Assert.AreEqual(expectedResult, actual, tolerance * 10);
        }

        [Theory]
        public void Atanh_ThrowsNumberExceptionWhenAbsOfInput1OrGreater([Range(1, 5, 0.2)] double input)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(string.Format(@"ATANH({0})", input.ToString(CultureInfo.InvariantCulture))));
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(string.Format(@"ATANH({0})", (-input).ToString(CultureInfo.InvariantCulture))));
        }

        [TestCase(0, 36, "0")]
        [TestCase(1, 36, "1")]
        [TestCase(2, 36, "2")]
        [TestCase(3, 36, "3")]
        [TestCase(4, 36, "4")]
        [TestCase(5, 36, "5")]
        [TestCase(6, 36, "6")]
        [TestCase(7, 36, "7")]
        [TestCase(8, 36, "8")]
        [TestCase(9, 36, "9")]
        [TestCase(10, 36, "A")]
        [TestCase(11, 36, "B")]
        [TestCase(12, 36, "C")]
        [TestCase(13, 36, "D")]
        [TestCase(14, 36, "E")]
        [TestCase(15, 36, "F")]
        [TestCase(16, 36, "G")]
        [TestCase(17, 36, "H")]
        [TestCase(18, 36, "I")]
        [TestCase(19, 36, "J")]
        [TestCase(20, 36, "K")]
        [TestCase(21, 36, "L")]
        [TestCase(22, 36, "M")]
        [TestCase(23, 36, "N")]
        [TestCase(24, 36, "O")]
        [TestCase(25, 36, "P")]
        [TestCase(26, 36, "Q")]
        [TestCase(27, 36, "R")]
        [TestCase(28, 36, "S")]
        [TestCase(29, 36, "T")]
        [TestCase(30, 36, "U")]
        [TestCase(31, 36, "V")]
        [TestCase(32, 36, "W")]
        [TestCase(33, 36, "X")]
        [TestCase(34, 36, "Y")]
        [TestCase(35, 36, "Z")]
        [TestCase(36, 36, "10")]
        [TestCase(255, 29, "8N")]
        [TestCase(255, 2, "11111111")]
        public void Base_ReturnsCorrectResultOnInput(int input, int theBase, string expectedResult)
        {
            var actual = (string)XLWorkbook.EvaluateExpr(string.Format(@"BASE({0}, {1})", input, theBase));
            Assert.AreEqual(expectedResult, actual);
        }

        [TestCase(255, 2, 3, "11111111")]
        [TestCase(255, 2, 8, "11111111")]
        [TestCase(255, 2, 10, "0011111111")]
        [TestCase(10, 3, 4, "0101")]
        public void Base_ReturnsCorrectResultOnInputWithMinimalLength(int input, int theBase, int minLength, string expectedResult)
        {
            var actual = (string)XLWorkbook.EvaluateExpr(string.Format(@"BASE({0}, {1}, {2})", input, theBase, minLength));
            Assert.AreEqual(expectedResult, actual);
        }

        [TestCase(@"""x""", "2", "2")]
        [TestCase("0", @"""x""", "2")]
        [TestCase("0", "2", @"""x""")]
        public void Base_ThrowsCellValueExceptionOnAnyInputNotANumber(string input, string theBase, string minLength)
        {
            Assert.AreEqual(XLCalculationErrorType.CellValue, XLWorkbook.EvaluateExpr($"BASE({input}, {theBase}, {minLength})"));
        }

        [Theory]
        public void Base_ThrowsNumberExceptionOnBaseSmallerThan2([Range(-2, 1)] int theBase)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(string.Format(@"BASE(0, {0})", theBase.ToString(CultureInfo.InvariantCulture))));
        }

        [Theory]
        public void Base_ThrowsNumberExceptionOnInputSmallerThan0([Range(-5, -1)] int input)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(string.Format(@"BASE({0}, 2)", input.ToString(CultureInfo.InvariantCulture))));
        }

        [Theory]
        public void Base_ThrowsNumberExceptionOnRadixGreaterThan36([Range(37, 40)] int radix)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(string.Format(@"BASE(1, {0})", radix.ToString(CultureInfo.InvariantCulture))));
        }

        [TestCase(24.3, 5, 25)]
        [TestCase(6.7, 1, 7)]
        [TestCase(-8.1, 2, -8)]
        [TestCase(5.5, 2.1, 6.3)]
        [TestCase(5.5, 0, 0)]
        [TestCase(-5.5, 2.1, -4.2)]
        [TestCase(-5.5, -2.1, -6.3)]
        [TestCase(-5.5, 0, 0)]
        public void Ceiling(double input, double significance, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"CEILING({input.ToInvariantString()}, {significance.ToInvariantString()})");
            Assert.AreEqual(expectedResult, actual, tolerance);
        }

        [TestCase(6.7, -1)]
        public void Ceiling_ThrowsNumberExceptionOnInvalidInput(double input, double significance)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr($"CEILING({input.ToInvariantString()}, {significance.ToInvariantString()})"));
        }

        [TestCase(24.3, 5, null, 25)]
        [TestCase(6.7, null, null, 7)]
        [TestCase(-8.1, 2, null, -8)]
        [TestCase(5.5, 2.1, 0, 6.3)]
        [TestCase(5.5, -2.1, 0, 6.3)]
        [TestCase(5.5, 0, 0, 0)]
        [TestCase(5.5, 2.1, -1, 6.3)]
        [TestCase(5.5, -2.1, -1, 6.3)]
        [TestCase(5.5, 0, -1, 0)]
        [TestCase(5.5, 2.1, 10, 6.3)]
        [TestCase(5.5, -2.1, 10, 6.3)]
        [TestCase(5.5, 0, 10, 0)]
        [TestCase(-5.5, 2.1, 0, -4.2)]
        [TestCase(-5.5, -2.1, 0, -4.2)]
        [TestCase(-5.5, 0, 0, 0)]
        [TestCase(-5.5, 2.1, -1, -6.3)]
        [TestCase(-5.5, -2.1, -1, -6.3)]
        [TestCase(-5.5, 0, -1, 0)]
        [TestCase(-5.5, 2.1, 10, -6.3)]
        [TestCase(-5.5, -2.1, 10, -6.3)]
        [TestCase(-5.5, 0, 10, 0)]
        public void CeilingMath(double input, double? step, int? mode, double expectedResult)
        {
            string parameters = input.ToString(CultureInfo.InvariantCulture);
            if (step != null)
            {
                parameters = parameters + ", " + step?.ToString(CultureInfo.InvariantCulture);
                if (mode != null)
                    parameters = parameters + ", " + mode?.ToString(CultureInfo.InvariantCulture);
            }

            var actual = (double)XLWorkbook.EvaluateExpr($"CEILING.MATH({parameters})");
            Assert.AreEqual(expectedResult, actual, tolerance);
        }

        [Theory]
        public void Combin_Returns1ForKis0OrKEqualsN([Range(0, 10)] int n)
        {
            var actual = XLWorkbook.EvaluateExpr(string.Format(@"COMBIN({0}, 0)", n));
            Assert.AreEqual(1, actual);

            var actual2 = XLWorkbook.EvaluateExpr(string.Format(@"COMBIN({0}, {0})", n));
            Assert.AreEqual(1, actual2);
        }

        [TestCase(4, 2, 6)]
        [TestCase(5, 2, 10)]
        [TestCase(6, 2, 15)]
        [TestCase(6, 3, 20)]
        [TestCase(7, 2, 21)]
        [TestCase(7, 3, 35)]
        public void Combin_ReturnsCorrectResults(int n, int k, int expectedResult)
        {
            var actual = XLWorkbook.EvaluateExpr(string.Format(@"COMBIN({0}, {1})", n, k));
            Assert.AreEqual(expectedResult, actual);

            var actual2 = XLWorkbook.EvaluateExpr(string.Format(@"COMBIN({0}, {1})", n, n - k));
            Assert.AreEqual(expectedResult, actual2);
        }

        [Theory]
        public void Combin_ReturnsNforKis1OrKisNminus1([Range(1, 10)] int n)
        {
            var actual = XLWorkbook.EvaluateExpr(string.Format(@"COMBIN({0}, 1)", n));
            Assert.AreEqual(n, actual);

            var actual2 = XLWorkbook.EvaluateExpr(string.Format(@"COMBIN({0}, {1})", n, n - 1));
            Assert.AreEqual(n, actual2);
        }

        [Theory]
        public void Combin_ThrowsNumberExceptionForAnyArgumentSmaller0([Range(-4, -1)] int smaller0)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(
                string.Format(
                    @"COMBIN({0}, {1})",
                    smaller0.ToString(CultureInfo.InvariantCulture),
                    (-smaller0).ToString(CultureInfo.InvariantCulture))));

            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(
                string.Format(
                    @"COMBIN({0}, {1})",
                    (-smaller0).ToString(CultureInfo.InvariantCulture),
                    smaller0.ToString(CultureInfo.InvariantCulture))));
        }

        [TestCase("\"no number\"")]
        [TestCase("\"\"")]
        public void Combin_ThrowsNumericExceptionForAnyArgumentNotNumeric(string input)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(
                string.Format(
                    @"COMBIN({0}, 1)",
                    input?.ToString(CultureInfo.InvariantCulture))));

            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(
                string.Format(
                    @"COMBIN(1, {0})",
                    input?.ToString(CultureInfo.InvariantCulture))));
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
        public void Combina_Returns1WhenChosenIs0([Range(0, 10)] int number)
        {
            Combina_CalculatesCorrectValues(number, 0, 1);
        }

        [TestCase(-1, 2)]
        [TestCase(-3, -2)]
        [TestCase(2, -2)]
        public void Combina_ThrowsNumExceptionOnInvalidValues(int number, int chosen)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(
                string.Format(
                    @"COMBINA({0}, {1})",
                    number.ToString(CultureInfo.InvariantCulture),
                    chosen.ToString(CultureInfo.InvariantCulture))));
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

        [TestCase(0, 1)]
        [TestCase(0.4, 0.921060994002885)]
        [TestCase(0.8, 0.696706709347165)]
        [TestCase(1.2, 0.362357754476674)]
        [TestCase(1.6, -0.0291995223012888)]
        [TestCase(2, -0.416146836547142)]
        [TestCase(2.4, -0.737393715541245)]
        [TestCase(2.8, -0.942222340668658)]
        [TestCase(3.2, -0.998294775794753)]
        [TestCase(3.6, -0.896758416334147)]
        [TestCase(4, -0.653643620863612)]
        [TestCase(4.4, -0.307332869978419)]
        [TestCase(4.8, 0.0874989834394464)]
        [TestCase(5.2, 0.468516671300377)]
        [TestCase(5.6, 0.77556587851025)]
        [TestCase(6, 0.960170286650366)]
        [TestCase(6.4, 0.993184918758193)]
        [TestCase(6.8, 0.869397490349825)]
        [TestCase(7.2, 0.608351314532255)]
        [TestCase(7.6, 0.251259842582256)]
        [TestCase(8, -0.145500033808614)]
        [TestCase(8.4, -0.519288654116686)]
        public void Cos_ReturnsCorrectResult(double input, double expectedResult)
        {
            var actualResult = (double)XLWorkbook.EvaluateExpr(string.Format("COS({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedResult, actualResult, tolerance);
        }

        [TestCase(0, 1)]
        [TestCase(0.4, 1.08107237183845)]
        [TestCase(0.8, 1.33743494630484)]
        [TestCase(1.2, 1.81065556732437)]
        [TestCase(1.6, 2.57746447119489)]
        [TestCase(2, 3.76219569108363)]
        [TestCase(2.4, 5.55694716696551)]
        [TestCase(2.8, 8.25272841686113)]
        [TestCase(3.2, 12.2866462005439)]
        [TestCase(3.6, 18.3127790830626)]
        [TestCase(4, 27.3082328360165)]
        [TestCase(4.4, 40.7315730024356)]
        [TestCase(4.8, 60.7593236328919)]
        [TestCase(5.2, 90.638879219786)]
        [TestCase(5.6, 135.215052644935)]
        [TestCase(6, 201.715636122456)]
        [TestCase(6.4, 300.923349714678)]
        [TestCase(6.8, 448.924202712783)]
        [TestCase(7.2, 669.715755490113)]
        [TestCase(7.6, 999.098197777775)]
        [TestCase(8, 1490.47916125218)]
        [TestCase(8.4, 2223.53348628359)]
        public void Cosh_ReturnsCorrectResult(double input, double expectedResult)
        {
            var actualResult = (double)XLWorkbook.EvaluateExpr(string.Format("COSH({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedResult, actualResult, tolerance);
            var actualResult2 = (double)XLWorkbook.EvaluateExpr(string.Format("COSH({0})", (-input).ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedResult, actualResult2, tolerance);
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
            Assert.AreEqual(XLCalculationErrorType.DivisionByZero, XLWorkbook.EvaluateExpr("COT(0)"));
        }

        [Test]
        public void Cot_On0_ThrowsDivisionByZeroException()
        {
            Assert.AreEqual(XLCalculationErrorType.DivisionByZero, XLWorkbook.EvaluateExpr(@"COTH(0)"));
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
        public void Csc_On0_ThrowsDivisionByZeroException()
        {
            Assert.AreEqual(XLCalculationErrorType.DivisionByZero, XLWorkbook.EvaluateExpr(@"CSC(0)"));
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
            Assert.AreEqual(XLCalculationErrorType.DivisionByZero, XLWorkbook.EvaluateExpr("CSCH(0)"));
        }

        [TestCase("FF", 16, 255)]
        [TestCase("111", 2, 7)]
        [TestCase("zap", 36, 45745)]
        public void Decimal(string inputString, int radix, int expectedResult)
        {
            var actualResult = XLWorkbook.EvaluateExpr($"DECIMAL(\"{inputString}\", {radix})");
            Assert.AreEqual(expectedResult, actualResult);
        }

        [Theory]
        public void Decimal_ReturnsErrorForRadiansGreater36([Range(37, 255)] int radix)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr($"DECIMAL(\"0\", {radix})"));
        }

        [Theory]
        public void Decimal_ReturnsErrorForRadiansSmaller2([Range(-5, 1)] int radix)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr($"DECIMAL(\"0\", {radix})"));
        }

        [Test]
        public void Decimal_ZeroIsZeroInAnyRadix([Range(2, 36)] int radix)
        {
            Assert.AreEqual(0, XLWorkbook.EvaluateExpr($"DECIMAL(\"0\", {radix})"));
        }

        [TestCase(0, 0)]
        [TestCase(Math.PI, 180)]
        [TestCase(Math.PI * 2, 360)]
        [TestCase(1, 57.2957795130823)]
        [TestCase(2, 114.591559026165)]
        [TestCase(3, 171.887338539247)]
        [TestCase(4, 229.183118052329)]
        [TestCase(5, 286.478897565412)]
        [TestCase(6, 343.774677078494)]
        [TestCase(7, 401.070456591576)]
        [TestCase(8, 458.366236104659)]
        [TestCase(9, 515.662015617741)]
        [TestCase(10, 572.957795130823)]
        [TestCase(Math.PI * 0.5, 90)]
        [TestCase(Math.PI * 1.5, 270)]
        [TestCase(Math.PI * 0.25, 45)]
        [TestCase(-1, -57.2957795130823)]
        public void Degrees_ReturnsCorrectResult(double input, double expected)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"DEGREES({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expected, actual, tolerance);
        }

        [TestCase(1.5, 2)]
        [TestCase(3, 4)]
        [TestCase(2, 2)]
        [TestCase(-1, -2)]
        [TestCase(0, 0)]
        [TestCase(Math.PI, 4)]
        public void Even_ReturnsCorrectResults(double input, int expectedResult)
        {
            var actual = (int)XLWorkbook.EvaluateExpr(string.Format(@"EVEN({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedResult, actual);
        }

        [TestCase(0, 1)]
        [TestCase(1, Math.E)]
        [TestCase(2, 7.38905609893065)]
        [TestCase(3, 20.0855369231877)]
        [TestCase(4, 54.5981500331442)]
        [TestCase(5, 148.413159102577)]
        [TestCase(6, 403.428793492735)]
        [TestCase(7, 1096.63315842846)]
        [TestCase(8, 2980.95798704173)]
        [TestCase(9, 8103.08392757538)]
        [TestCase(10, 22026.4657948067)]
        [TestCase(11, 59874.1417151978)]
        [TestCase(12, 162754.791419004)]
        public void Exp_ReturnsCorrectResults(double input, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"EXP({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedResult, actual, tolerance);
        }

        [TestCase(0, 1L)]
        [TestCase(1, 1L)]
        [TestCase(2, 2L)]
        [TestCase(3, 6L)]
        [TestCase(4, 24L)]
        [TestCase(5, 120L)]
        [TestCase(6, 720L)]
        [TestCase(7, 5040L)]
        [TestCase(8, 40320L)]
        [TestCase(9, 362880L)]
        [TestCase(10, 3628800L)]
        [TestCase(11, 39916800L)]
        [TestCase(12, 479001600L)]
        [TestCase(13, 6227020800L)]
        [TestCase(14, 87178291200L)]
        [TestCase(15, 1307674368000L)]
        [TestCase(16, 20922789888000L)]
        [TestCase(0.1, 1L)]
        [TestCase(2.3, 2L)]
        [TestCase(2.8, 2L)]
        public void Fact_ReturnsCorrectResult(double input, long expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"FACT({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedResult, actual);
        }

        [Theory]
        public void Fact_ThrowsNumberExceptionForNegativeInput([Range(-10, -1)] int input)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(string.Format(@"FACT({0})", input.ToString(CultureInfo.InvariantCulture))));
        }

        [Test]
        public void Fact_ThrowsValueExceptionForNonNumericInput()
        {
            Assert.AreEqual(XLCalculationErrorType.CellValue, XLWorkbook.EvaluateExpr(string.Format(@"FACT(""x"")")));
        }

        [TestCase(0, 1L)]
        [TestCase(1, 1L)]
        [TestCase(2, 2L)]
        [TestCase(3, 3L)]
        [TestCase(4, 8L)]
        [TestCase(5, 15L)]
        [TestCase(6, 48L)]
        [TestCase(7, 105L)]
        [TestCase(8, 384L)]
        [TestCase(9, 945L)]
        [TestCase(10, 3840L)]
        [TestCase(11, 10395L)]
        [TestCase(12, 46080L)]
        [TestCase(13, 135135L)]
        [TestCase(14, 645120)]
        [TestCase(15, 2027025)]
        [TestCase(16, 10321920)]
        [TestCase(-1, 1L)]
        [TestCase(0, 1)]
        [TestCase(0.1, 1L)]
        [TestCase(1.4, 1L)]
        [TestCase(2.3, 2L)]
        [TestCase(2.8, 2L)]
        public void FactDouble_ReturnsCorrectResult(double input, long expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"FACTDOUBLE({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expectedResult, actual);
        }

        [Theory]
        public void FactDouble_ThrowsNumberExceptionForInputSmallerThanMinus1([Range(-10, -2)] int input)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr(string.Format(@"FACTDOUBLE({0})", input.ToString(CultureInfo.InvariantCulture))));
        }

        [Test]
        public void FactDouble_ThrowsValueExceptionForNonNumericInput()
        {
            Assert.AreEqual(XLCalculationErrorType.CellValue, XLWorkbook.EvaluateExpr(string.Format(@"FACTDOUBLE(""x"")")));
        }

        [TestCase(24.3, 5, 20)]
        [TestCase(6.7, 1, 6)]
        [TestCase(-8.1, 2, -10)]
        [TestCase(5.5, 2.1, 4.2)]
        [TestCase(-5.5, 2.1, -6.3)]
        [TestCase(-5.5, -2.1, -4.2)]
        public void Floor(double input, double significance, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"FLOOR({input.ToInvariantString()}, {significance.ToInvariantString()})");
            Assert.AreEqual(expectedResult, actual, tolerance);
        }

        [TestCase(6.7, 0)]
        [TestCase(-6.7, 0)]
        public void Floor_ThrowsDivisionByZeroOnZeroSignificance(double input, double significance)
        {
            Assert.AreEqual(XLCalculationErrorType.DivisionByZero, XLWorkbook.EvaluateExpr($"FLOOR({input.ToInvariantString()}, {significance.ToInvariantString()})"));
        }

        [TestCase(6.7, -1)]
        public void Floor_ThrowsNumberExceptionOnInvalidInput(double input, double significance)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr($"FLOOR({input.ToInvariantString()}, {significance.ToInvariantString()})"));
        }

        [Test]
        // Functions have to support a period first before we can implement this
        [TestCase(24.3, 5, null, 20)]
        [TestCase(6.7, null, null, 6)]
        [TestCase(-8.1, 2, null, -10)]
        [TestCase(5.5, 2.1, 0, 4.2)]
        [TestCase(5.5, -2.1, 0, 4.2)]
        [TestCase(5.5, 0, 0, 0)]
        [TestCase(5.5, 2.1, -1, 4.2)]
        [TestCase(5.5, -2.1, -1, 4.2)]
        [TestCase(5.5, 0, -2, 0)]
        [TestCase(5.5, 2.1, 10, 4.2)]
        [TestCase(5.5, -2.1, 10, 4.2)]
        [TestCase(5.5, 0, 10, 0)]
        [TestCase(-5.5, 2.1, 0, -6.3)]
        [TestCase(-5.5, -2.1, 0, -6.3)]
        [TestCase(-5.5, 0, 0, 0)]
        [TestCase(-5.5, 2.1, -1, -4.2)]
        [TestCase(-5.5, -2.1, -1, -4.2)]
        [TestCase(-5.5, 0, -1, 0)]
        [TestCase(-5.5, 2.1, 10, -4.2)]
        [TestCase(-5.5, -2.1, 10, -4.2)]
        [TestCase(-5.5, 0, 0, 0)]
        public void FloorMath(double input, double? step, int? mode, double expectedResult)
        {
            string parameters = input.ToString(CultureInfo.InvariantCulture);
            if (step != null)
            {
                parameters = parameters + ", " + step?.ToString(CultureInfo.InvariantCulture);
                if (mode != null)
                    parameters = parameters + ", " + mode?.ToString(CultureInfo.InvariantCulture);
            }

            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"FLOOR.MATH({0})", parameters));
            Assert.AreEqual(expectedResult, actual, tolerance);
        }

        [TestCase(8.9, 8)]
        [TestCase(-8.9, -9)]
        public void Int(double input, double expected)
        {
            var actual = XLWorkbook.EvaluateExpr(string.Format(@"INT({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(expected, actual);
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

        [TestCase(10, 3, ExpectedResult = 9.0)]
        [TestCase(10.5, 3, ExpectedResult = 12.0)]
        [TestCase(10.4, 3, ExpectedResult = 9.0)]
        [TestCase(-10, -3, ExpectedResult = -9.0)]
        [TestCase(1.3, 0.2, ExpectedResult = 1.4)]
        [TestCase(5677.912288, 10, ExpectedResult = 5680.0)]
        [TestCase(5674.912288, 10, ExpectedResult = 5670.0)]
        [TestCase(0.5, 1, ExpectedResult = 1.0)]
        [TestCase(0.49999, 1, ExpectedResult = 0.0)]
        [TestCase(0.5, 1, ExpectedResult = 1.0)]
        [TestCase(0.49999, 1, ExpectedResult = 0.0)]
        [TestCase(0.5, 1, ExpectedResult = 1.0)]
        [TestCase(0.49999, 1, ExpectedResult = 0.0)]
        [TestCase(-13.4, -3, ExpectedResult = -12.0)]
        [TestCase(-13.5, -3, ExpectedResult = -15.0)]
        [TestCase(0.9, 0.2, ExpectedResult = 1.0)]
        [TestCase(0.89999, 0.2, ExpectedResult = 0.8)]
        [TestCase(15.5, 3, ExpectedResult = 15.0)]
        [TestCase(1.4, 0.5, ExpectedResult = 1.5)]
        [DefaultFloatingPointTolerance(1e-12)]
        public double MRound(double number, double multiple)
        {
            return (double)XLWorkbook.EvaluateExpr($"MROUND({number}, {multiple})");
        }

        [TestCase(123456.123, -10)]
        [TestCase(-123456.123, 5)]
        public void MRoundExceptions(double number, double multiple)
        {
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr($"MROUND({number}, {multiple})"));
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
            Assert.AreEqual(XLCalculationErrorType.CellValue, XLWorkbook.EvaluateExpr(@"SEC(""number"")"));
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

        [Test]
        public void SumIf_ReturnsCorrectValues_WhenCalledOnFullColumn()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Data");
                var data = new object[]
                {
                    new { Id = "A", Value = 2},
                    new { Id = "B", Value = 3},
                    new { Id = "C", Value = 2},
                    new { Id = "A", Value = 1},
                    new { Id = "B", Value = 4}
                };
                ws.Cell("A1").InsertTable(data);
                var formula = "=SUMIF(A:A,\"=A\",B:B)";
                var value = ws.Evaluate(formula);
                Assert.AreEqual(3, value);
            }
        }

        [Test]
        public void SumIf_ReturnsCorrectValues_WhenFormulaBelongToSameRange()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Data");
                var data = new object[]
                {
                    new { Id = "A", Value = 2},
                    new { Id = "B", Value = 3},
                    new { Id = "C", Value = 2},
                    new { Id = "A", Value = 1},
                    new { Id = "B", Value = 4},
                };
                ws.Cell("A1").InsertTable(data);
                ws.Cell("A7").SetValue("Sum A");
                // SUMIF formula
                var formula = "=SUMIF(A:A,\"=A\",B:B)";
                ws.Cell("B7").SetFormulaA1(formula);
                var value = ws.Cell("B7").Value;
                Assert.AreEqual(3, value);
            }
        }

        /// <summary>
        /// refers to Example 2 to SumIf from the Excel documentation.
        /// As SumIfs should behave the same if called with three parameters, we can take that example here again.
        /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
        /// </summary>
        /// <param name="expectedResult"></param>
        /// <param name="formula"></param>
        [TestCase(2000, "SUMIFS(C2:C7, A2:A7, \"Fruits\")")]
        [TestCase(12000, "SUMIFS(C2:C7, A2:A7, \"Vegetables\")")]
        [TestCase(4300, "SUMIFS(C2:C7, B2:B7, \"*es\")")]
        [TestCase(400, "SUMIFS(C2:C7, A2:A7, \"\")")]
        public void SumIfs_ReturnsCorrectValues_ReferenceExample2FromMicrosoft(int expectedResult, string formula)
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

                var actualResult = ws.Evaluate(formula).CastTo<double>();
                Assert.AreEqual(expectedResult, actualResult);
            }
        }

        /// <summary>
        /// refers to Example 1 to SumIf from the Excel documentation.
        /// As SumIfs should behave the same if called with three parameters, but in a different order
        /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
        /// </summary>
        /// <param name="expectedOutcome"></param>
        /// <param name="formula"></param>
        [TestCase(63000, "SUMIFS(B1:B4, A1:A4, \">160000\")")]
        [TestCase(21000, "SUMIFS(B1:B4, A1:A4, 300000)")]
        [TestCase(28000, "SUMIFS(B1:B4, A1:A4, \">\" &C1)")]
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
        /// refers to example data and formula to SumIfs in the Excel documentation,
        /// <see cref="https://support.office.com/en-us/article/SUMIFS-function-c9e748f5-7ea7-455d-9406-611cebce642b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
        /// </summary>
        [TestCase(20, "=SUMIFS(A2:A9, B2:B9, \"=A*\", C2:C9, \"Tom\")")]
        [TestCase(30, "=SUMIFS(A2:A9, B2:B9, \"<>Bananas\", C2:C9, \"Tom\")")]
        public void SumIfs_ReturnsCorrectValues_ReferenceExampleFromMicrosoft(
            int expectedResult,
            string formula)
        {
            using (var wb = new XLWorkbook())
            {
                wb.ReferenceStyle = XLReferenceStyle.A1;
                var ws = wb.AddWorksheet("Sheet1");

                var row = 2;

                ws.Cell(row, 1).Value = 5;
                ws.Cell(row, 2).Value = "Apples";
                ws.Cell(row, 3).Value = "Tom";
                row++;

                ws.Cell(row, 1).Value = 4;
                ws.Cell(row, 2).Value = "Apples";
                ws.Cell(row, 3).Value = "Sarah";
                row++;

                ws.Cell(row, 1).Value = 15;
                ws.Cell(row, 2).Value = "Artichokes";
                ws.Cell(row, 3).Value = "Tom";
                row++;

                ws.Cell(row, 1).Value = 3;
                ws.Cell(row, 2).Value = "Artichokes";
                ws.Cell(row, 3).Value = "Sarah";
                row++;

                ws.Cell(row, 1).Value = 22;
                ws.Cell(row, 2).Value = "Bananas";
                ws.Cell(row, 3).Value = "Tom";
                row++;

                ws.Cell(row, 1).Value = 12;
                ws.Cell(row, 2).Value = "Bananas";
                ws.Cell(row, 3).Value = "Sarah";
                row++;

                ws.Cell(row, 1).Value = 10;
                ws.Cell(row, 2).Value = "Carrots";
                ws.Cell(row, 3).Value = "Tom";
                row++;

                ws.Cell(row, 1).Value = 33;
                ws.Cell(row, 2).Value = "Carrots";
                ws.Cell(row, 3).Value = "Sarah";

                var actualResult = ws.Evaluate(formula).CastTo<Double>();

                Assert.AreEqual(expectedResult, actualResult, tolerance);
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

                Assert.AreEqual(XLCalculationErrorType.NoValueAvailable, ws.Evaluate("SUMPRODUCT(A1:A10, B1:B5)"));

                // Blank cells and cells with text should be treated as zeros
                ws.Range("A1:A5").Clear();
                Assert.AreEqual(110, ws.Evaluate("SUMPRODUCT(A1:A10, B1:B10)"));

                ws.Range("A1:A5").SetValue("asdf");
                Assert.AreEqual(110, ws.Evaluate("SUMPRODUCT(A1:A10, B1:B10)"));
            }
        }

        [Test]
        public void Trunc()
        {
            var input = 27.64799257;
            var expectedResult = 27;
            var actual = (double)XLWorkbook.EvaluateExpr($"TRUNC({input.ToString(CultureInfo.InvariantCulture)})");
            Assert.AreEqual(expectedResult, actual);
        }

        [TestCase(27.64799257, -1, 20)]
        [TestCase(27.64799257, 0, 27)]
        [TestCase(27.64799257, 1, 27.6)]
        [TestCase(27.64799257, 4, 27.6479)]
        public void Trunc_Specify_Digits(double input, int digits, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"TRUNC({input.ToString(CultureInfo.InvariantCulture)}, {digits})");
            Assert.AreEqual(expectedResult, actual);
        }
    }
}
