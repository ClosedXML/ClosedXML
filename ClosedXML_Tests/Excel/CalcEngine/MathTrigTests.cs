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
