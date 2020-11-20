using ClosedXML.Excel;
using NUnit.Framework;
using System;

namespace ClosedXML_Tests.Excel.CalcEngine
{
    [TestFixture]
    public class FunctionsTests
    {
        [SetUp]
        public void Init()
        {
            // Make sure tests run on a deterministic culture
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        }

        [Test]
        public void Asc()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Asc(""Text"")");
            Assert.AreEqual("Text", actual);
        }

        [Test]
        public void Clean()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(String.Format(@"Clean(""A{0}B"")", Environment.NewLine));
            Assert.AreEqual("AB", actual);
        }

        [Test]
        public void Combin()
        {
            object actual1 = XLWorkbook.EvaluateExpr("Combin(200, 2)");
            Assert.AreEqual(19900.0, actual1);

            object actual2 = XLWorkbook.EvaluateExpr("Combin(20.1, 2.9)");
            Assert.AreEqual(190.0, actual2);
        }

        [Test]
        public void Degrees()
        {
            object actual1 = XLWorkbook.EvaluateExpr("Degrees(180)");
            Assert.IsTrue(Math.PI - (double)actual1 < XLHelper.Epsilon);
        }

        [Test]
        public void Dollar()
        {
            object actual = XLWorkbook.EvaluateExpr("Dollar(12345.123)");
            Assert.AreEqual(TestHelper.CurrencySymbol + "12,345.12", actual);

            actual = XLWorkbook.EvaluateExpr("Dollar(12345.123, 1)");
            Assert.AreEqual(TestHelper.CurrencySymbol + "12,345.1", actual);
        }

        [Test]
        public void Even()
        {
            object actual = XLWorkbook.EvaluateExpr("Even(3)");
            Assert.AreEqual(4, actual);

            actual = XLWorkbook.EvaluateExpr("Even(2)");
            Assert.AreEqual(2, actual);

            actual = XLWorkbook.EvaluateExpr("Even(-1)");
            Assert.AreEqual(-2, actual);

            actual = XLWorkbook.EvaluateExpr("Even(-2)");
            Assert.AreEqual(-2, actual);

            actual = XLWorkbook.EvaluateExpr("Even(0)");
            Assert.AreEqual(0, actual);

            actual = XLWorkbook.EvaluateExpr("Even(1.5)");
            Assert.AreEqual(2, actual);

            actual = XLWorkbook.EvaluateExpr("Even(2.01)");
            Assert.AreEqual(4, actual);
        }

        [Test]
        public void Exact()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr("Exact(\"A\", \"A\")");
            Assert.AreEqual(true, actual);

            actual = XLWorkbook.EvaluateExpr("Exact(\"A\", \"a\")");
            Assert.AreEqual(false, actual);
        }

        [Test]
        public void Fact()
        {
            object actual = XLWorkbook.EvaluateExpr("Fact(5.9)");
            Assert.AreEqual(120.0, actual);
        }

        [Test]
        public void FactDouble()
        {
            object actual1 = XLWorkbook.EvaluateExpr("FactDouble(6)");
            Assert.AreEqual(48.0, actual1);
            object actual2 = XLWorkbook.EvaluateExpr("FactDouble(7)");
            Assert.AreEqual(105.0, actual2);
        }

        [Test]
        public void Fixed()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr("Fixed(12345.123)");
            Assert.AreEqual("12,345.12", actual);

            actual = XLWorkbook.EvaluateExpr("Fixed(12345.123, 1)");
            Assert.AreEqual("12,345.1", actual);

            actual = XLWorkbook.EvaluateExpr("Fixed(12345.123, 1, TRUE)");
            Assert.AreEqual("12345.1", actual);
        }

        [Test]
        public void Formula_from_another_sheet()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws1 = wb.AddWorksheet("ws1");
            ws1.FirstCell().SetValue(1).CellRight().SetFormulaA1("A1 + 1");
            IXLWorksheet ws2 = wb.AddWorksheet("ws2");
            ws2.FirstCell().SetFormulaA1("ws1!B1 + 1");
            object v = ws2.FirstCell().Value;
            Assert.AreEqual(3.0, v);
        }

        [Test]
        public void Gcd()
        {
            object actual = XLWorkbook.EvaluateExpr("Gcd(24, 36)");
            Assert.AreEqual(12, actual);

            object actual1 = XLWorkbook.EvaluateExpr("Gcd(5, 0)");
            Assert.AreEqual(5, actual1);

            object actual2 = XLWorkbook.EvaluateExpr("Gcd(0, 5)");
            Assert.AreEqual(5, actual2);

            object actual3 = XLWorkbook.EvaluateExpr("Gcd(240, 360, 30)");
            Assert.AreEqual(30, actual3);
        }

        [Test]
        public void Lcm()
        {
            object actual = XLWorkbook.EvaluateExpr("Lcm(24, 36)");
            Assert.AreEqual(72, actual);

            object actual1 = XLWorkbook.EvaluateExpr("Lcm(5, 0)");
            Assert.AreEqual(0, actual1);

            object actual2 = XLWorkbook.EvaluateExpr("Lcm(0, 5)");
            Assert.AreEqual(0, actual2);

            object actual3 = XLWorkbook.EvaluateExpr("Lcm(240, 360, 30)");
            Assert.AreEqual(720, actual3);
        }

        [Test]
        public void MDetem()
        {
            IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");
            ws.Cell("A1").SetValue(2).CellRight().SetValue(4);
            ws.Cell("A2").SetValue(3).CellRight().SetValue(5);

            Object actual;

            ws.Cell("A5").FormulaA1 = "MDeterm(A1:B2)";
            actual = ws.Cell("A5").Value;

            Assert.IsTrue(XLHelper.AreEqual(-2.0, (double)actual));

            ws.Cell("A6").FormulaA1 = "Sum(A5)";
            actual = ws.Cell("A6").Value;

            Assert.IsTrue(XLHelper.AreEqual(-2.0, (double)actual));

            ws.Cell("A7").FormulaA1 = "Sum(MDeterm(A1:B2))";
            actual = ws.Cell("A7").Value;

            Assert.IsTrue(XLHelper.AreEqual(-2.0, (double)actual));
        }

        [Test]
        public void MInverse()
        {
            IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");
            ws.Cell("A1").SetValue(1).CellRight().SetValue(2).CellRight().SetValue(1);
            ws.Cell("A2").SetValue(3).CellRight().SetValue(4).CellRight().SetValue(-1);
            ws.Cell("A3").SetValue(0).CellRight().SetValue(2).CellRight().SetValue(0);

            Object actual;

            ws.Cell("A5").FormulaA1 = "MInverse(A1:C3)";
            actual = ws.Cell("A5").Value;

            Assert.IsTrue(XLHelper.AreEqual(0.25, (double)actual));

            ws.Cell("A6").FormulaA1 = "Sum(A5)";
            actual = ws.Cell("A6").Value;

            Assert.IsTrue(XLHelper.AreEqual(0.25, (double)actual));

            ws.Cell("A7").FormulaA1 = "Sum(MInverse(A1:C3))";
            actual = ws.Cell("A7").Value;

            Assert.IsTrue(XLHelper.AreEqual(0.5, (double)actual));
        }

        [Test]
        public void MMult()
        {
            IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");
            ws.Cell("A1").SetValue(2).CellRight().SetValue(4);
            ws.Cell("A2").SetValue(3).CellRight().SetValue(5);
            ws.Cell("A3").SetValue(2).CellRight().SetValue(4);
            ws.Cell("A4").SetValue(3).CellRight().SetValue(5);

            Object actual;

            ws.Cell("A5").FormulaA1 = "MMult(A1:B2, A3:B4)";
            actual = ws.Cell("A5").Value;

            Assert.AreEqual(16.0, actual);

            ws.Cell("A6").FormulaA1 = "Sum(A5)";
            actual = ws.Cell("A6").Value;

            Assert.AreEqual(16.0, actual);

            ws.Cell("A7").FormulaA1 = "Sum(MMult(A1:B2, A3:B4))";
            actual = ws.Cell("A7").Value;

            Assert.AreEqual(102.0, actual);
        }

        [Test]
        public void Mod()
        {
            object actual = XLWorkbook.EvaluateExpr("Mod(3, 2)");
            Assert.AreEqual(1, actual);

            object actual1 = XLWorkbook.EvaluateExpr("Mod(-3, 2)");
            Assert.AreEqual(1, actual1);

            object actual2 = XLWorkbook.EvaluateExpr("Mod(3, -2)");
            Assert.AreEqual(-1, actual2);

            object actual3 = XLWorkbook.EvaluateExpr("Mod(-3, -2)");
            Assert.AreEqual(-1, actual3);
        }

        [Test]
        public void Multinomial()
        {
            object actual = XLWorkbook.EvaluateExpr("Multinomial(2,3,4)");
            Assert.AreEqual(1260.0, actual);
        }

        [Test]
        public void Odd()
        {
            object actual = XLWorkbook.EvaluateExpr("Odd(1.5)");
            Assert.AreEqual(3, actual);

            object actual1 = XLWorkbook.EvaluateExpr("Odd(3)");
            Assert.AreEqual(3, actual1);

            object actual2 = XLWorkbook.EvaluateExpr("Odd(2)");
            Assert.AreEqual(3, actual2);

            object actual3 = XLWorkbook.EvaluateExpr("Odd(-1)");
            Assert.AreEqual(-1, actual3);

            object actual4 = XLWorkbook.EvaluateExpr("Odd(-2)");
            Assert.AreEqual(-3, actual4);

            actual = XLWorkbook.EvaluateExpr("Odd(0)");
            Assert.AreEqual(1, actual);
        }

        [Test]
        public void Product()
        {
            object actual = XLWorkbook.EvaluateExpr("Product(2,3,4)");
            Assert.AreEqual(24.0, actual);
        }

        [Test]
        public void Quotient()
        {
            object actual = XLWorkbook.EvaluateExpr("Quotient(5,2)");
            Assert.AreEqual(2, actual);

            actual = XLWorkbook.EvaluateExpr("Quotient(4.5,3.1)");
            Assert.AreEqual(1, actual);

            actual = XLWorkbook.EvaluateExpr("Quotient(-10,3)");
            Assert.AreEqual(-3, actual);
        }

        [Test]
        public void Radians()
        {
            object actual = XLWorkbook.EvaluateExpr("Radians(270)");
            Assert.IsTrue(Math.Abs(4.71238898038469 - (double)actual) < XLHelper.Epsilon);
        }

        [Test]
        public void Roman()
        {
            object actual = XLWorkbook.EvaluateExpr("Roman(3046, 1)");
            Assert.AreEqual("MMMXLVI", actual);

            actual = XLWorkbook.EvaluateExpr("Roman(270)");
            Assert.AreEqual("CCLXX", actual);

            actual = XLWorkbook.EvaluateExpr("Roman(3999, true)");
            Assert.AreEqual("MMMCMXCIX", actual);
        }

        [Test]
        public void Round()
        {
            object actual = XLWorkbook.EvaluateExpr("Round(2.15, 1)");
            Assert.AreEqual(2.2, actual);

            actual = XLWorkbook.EvaluateExpr("Round(2.149, 1)");
            Assert.AreEqual(2.1, actual);

            actual = XLWorkbook.EvaluateExpr("Round(-1.475, 2)");
            Assert.AreEqual(-1.48, actual);

            actual = XLWorkbook.EvaluateExpr("Round(21.5, -1)");
            Assert.AreEqual(20.0, actual);

            actual = XLWorkbook.EvaluateExpr("Round(626.3, -3)");
            Assert.AreEqual(1000.0, actual);

            actual = XLWorkbook.EvaluateExpr("Round(1.98, -1)");
            Assert.AreEqual(0.0, actual);

            actual = XLWorkbook.EvaluateExpr("Round(-50.55, -2)");
            Assert.AreEqual(-100.0, actual);

            actual = XLWorkbook.EvaluateExpr("ROUND(59 * 0.535, 2)"); // (59 * 0.535) = 31.565
            Assert.AreEqual(31.57, actual);

            actual = XLWorkbook.EvaluateExpr("ROUND(59 * -0.535, 2)"); // (59 * -0.535) = -31.565
            Assert.AreEqual(-31.57, actual);
        }

        [Test]
        public void RoundDown()
        {
            object actual = XLWorkbook.EvaluateExpr("RoundDown(3.2, 0)");
            Assert.AreEqual(3.0, actual);

            actual = XLWorkbook.EvaluateExpr("RoundDown(76.9, 0)");
            Assert.AreEqual(76.0, actual);

            actual = XLWorkbook.EvaluateExpr("RoundDown(3.14159, 3)");
            Assert.AreEqual(3.141, actual);

            actual = XLWorkbook.EvaluateExpr("RoundDown(-3.14159, 1)");
            Assert.AreEqual(-3.1, actual);

            actual = XLWorkbook.EvaluateExpr("RoundDown(31415.92654, -2)");
            Assert.AreEqual(31400.0, actual);

            actual = XLWorkbook.EvaluateExpr("RoundDown(0, 3)");
            Assert.AreEqual(0.0, actual);
        }

        [Test]
        public void RoundUp()
        {
            object actual = XLWorkbook.EvaluateExpr("RoundUp(3.2, 0)");
            Assert.AreEqual(4.0, actual);

            actual = XLWorkbook.EvaluateExpr("RoundUp(76.9, 0)");
            Assert.AreEqual(77.0, actual);

            actual = XLWorkbook.EvaluateExpr("RoundUp(3.14159, 3)");
            Assert.AreEqual(3.142, actual);

            actual = XLWorkbook.EvaluateExpr("RoundUp(-3.14159, 1)");
            Assert.AreEqual(-3.2, actual);

            actual = XLWorkbook.EvaluateExpr("RoundUp(31415.92654, -2)");
            Assert.AreEqual(31500.0, actual);

            actual = XLWorkbook.EvaluateExpr("RoundUp(0, 3)");
            Assert.AreEqual(0.0, actual);
        }

        [Test]
        public void SeriesSum()
        {
            object actual = XLWorkbook.EvaluateExpr("SERIESSUM(2,3,4,5)");
            Assert.AreEqual(40.0, actual);

            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A2").FormulaA1 = "PI()/4";
            ws.Cell("A3").Value = 1;
            ws.Cell("A4").FormulaA1 = "-1/FACT(2)";
            ws.Cell("A5").FormulaA1 = "1/FACT(4)";
            ws.Cell("A6").FormulaA1 = "-1/FACT(6)";

            actual = ws.Evaluate("SERIESSUM(A2,0,2,A3:A6)");
            Assert.IsTrue(Math.Abs(0.70710321482284566 - (double)actual) < XLHelper.Epsilon);
        }

        [Test]
        public void SqrtPi()
        {
            object actual = XLWorkbook.EvaluateExpr("SqrtPi(1)");
            Assert.IsTrue(Math.Abs(1.7724538509055159 - (double)actual) < XLHelper.Epsilon);

            actual = XLWorkbook.EvaluateExpr("SqrtPi(2)");
            Assert.IsTrue(Math.Abs(2.5066282746310002 - (double)actual) < XLHelper.Epsilon);
        }

        [Test]
        public void SubtotalAverage()
        {
            object actual = XLWorkbook.EvaluateExpr("Subtotal(1,2,3)");
            Assert.AreEqual(2.5, actual);

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(1,""A"",3, 2)");
            Assert.AreEqual(2.5, actual);
        }

        [Test]
        public void SubtotalCount()
        {
            object actual = XLWorkbook.EvaluateExpr("Subtotal(2,2,3)");
            Assert.AreEqual(2, actual);

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(2,""A"",3)");
            Assert.AreEqual(1, actual);
        }

        [Test]
        public void SubtotalCountA()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr("Subtotal(3,2,3)");
            Assert.AreEqual(2.0, actual);

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(3,"""",3)");
            Assert.AreEqual(1.0, actual);
        }

        [Test]
        public void SubtotalMax()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(4,2,3,""A"")");
            Assert.AreEqual(3.0, actual);
        }

        [Test]
        public void SubtotalMin()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(5,2,3,""A"")");
            Assert.AreEqual(2.0, actual);
        }

        [Test]
        public void SubtotalProduct()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(6,2,3,""A"")");
            Assert.AreEqual(6.0, actual);
        }

        [Test]
        public void SubtotalStDev()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(7,2,3,""A"")");
            Assert.IsTrue(Math.Abs(0.70710678118654757 - (double)actual) < XLHelper.Epsilon);
        }

        [Test]
        public void SubtotalStDevP()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(8,2,3,""A"")");
            Assert.AreEqual(0.5, actual);
        }

        [Test]
        public void SubtotalSum()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(9,2,3,""A"")");
            Assert.AreEqual(5.0, actual);
        }

        [Test]
        public void SubtotalVar()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(10,2,3,""A"")");
            Assert.IsTrue(Math.Abs(0.5 - (double)actual) < XLHelper.Epsilon);
        }

        [Test]
        public void SubtotalVarP()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(11,2,3,""A"")");
            Assert.AreEqual(0.25, actual);
        }

        [Test]
        public void SubtotalCalc()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.NamedRanges.Add("subtotalrange", "A37:A38");

            ws.Cell("A1").Value = 2;
            ws.Cell("A2").Value = 4;
            ws.Cell("A3").FormulaA1 = "SUBTOTAL(9, A1:A2)"; // simple add subtotal
            ws.Cell("A4").Value = 8;
            ws.Cell("A5").Value = 16;
            ws.Cell("A6").FormulaA1 = "SUBTOTAL(9, A4:A5)"; // simple add subtotal
            ws.Cell("A7").Value = 32;
            ws.Cell("A8").Value = 64;
            ws.Cell("A9").FormulaA1 = "SUM(A7:A8)"; // func but not subtotal
            ws.Cell("A10").Value = 128;
            ws.Cell("A11").Value = 256;
            ws.Cell("A12").FormulaA1 = "SUBTOTAL(1, A10:A11)"; // simple avg subtotal
            ws.Cell("A13").Value = 512;
            ws.Cell("A14").FormulaA1 = "SUBTOTAL(9, A1:A13)"; // subtotals in range
            ws.Cell("A15").Value = 1024;
            ws.Cell("A16").Value = 2048;
            ws.Cell("A17").FormulaA1 = "42 + SUBTOTAL(9, A15:A16)"; // simple add subtotal in formula
            ws.Cell("A18").Value = 4096;
            ws.Cell("A19").FormulaA1 = "SUBTOTAL(9, A15:A18)"; // subtotals in range
            ws.Cell("A20").Value = 8192;
            ws.Cell("A21").Value = 16384;
            ws.Cell("A22").FormulaA1 = @"32768 * SEARCH(""SUBTOTAL(9, A1:A2)"", A28)"; // subtotal literal in formula
            ws.Cell("A23").FormulaA1 = "SUBTOTAL(9, A20:A22)"; // subtotal literal in formula in range
            ws.Cell("A24").Value = 65536;
            ws.Cell("A25").FormulaA1 = "A23"; // link to subtotal
            ws.Cell("A26").FormulaA1 = "PRODUCT(SUBTOTAL(9, A24:A25), 2)"; // subtotal as parameter in func
            ws.Cell("A27").Value = 131072;
            ws.Cell("A28").Value = "SUBTOTAL(9, A1:A2)"; // subtotal literal
            ws.Cell("A29").FormulaA1 = "SUBTOTAL(9, A27:A28)"; // subtotal literal in range
            ws.Cell("A30").FormulaA1 = "SUBTOTAL(9, A31:A32)"; // simple add subtotal backward
            ws.Cell("A31").Value = 262144;
            ws.Cell("A32").Value = 524288;
            ws.Cell("A33").FormulaA1 = "SUBTOTAL(9, A20:A32)"; // subtotals in range
            ws.Cell("A34").FormulaA1 = @"SUBTOTAL(VALUE(""9""), A1:A33, A35:A41)"; // func as parameter in subtotal and many ranges
            ws.Cell("A35").Value = 1048576;
            ws.Cell("A36").FormulaA1 = "SUBTOTAL(9, A31:A32, A35)"; // many ranges
            ws.Cell("A37").Value = 2097152;
            ws.Cell("A38").Value = 4194304;
            ws.Cell("A39").FormulaA1 = "SUBTOTAL(3*3, subtotalrange)"; // formula as parameter in subtotal and named range
            ws.Cell("A40").Value = 8388608;
            ws.Cell("A41").FormulaA1 = "PRODUCT(SUBTOTAL(A4+1, A35:A40), 2)"; // formula with link as parameter in subtotal
            ws.Cell("A42").FormulaA1 = "PRODUCT(SUBTOTAL(A4+1, A35:A40), 2) + SUBTOTAL(A4+1, A35:A40)"; // two subtotals in one formula

            Assert.AreEqual(6, ws.Cell("A3").Value);
            Assert.AreEqual(24, ws.Cell("A6").Value);
            Assert.AreEqual(192, ws.Cell("A12").Value);
            Assert.AreEqual(1118, ws.Cell("A14").Value);
            Assert.AreEqual(3114, ws.Cell("A17").Value);
            Assert.AreEqual(7168, ws.Cell("A19").Value);
            Assert.AreEqual(57344, ws.Cell("A23").Value);
            Assert.AreEqual(245760, ws.Cell("A26").Value);
            Assert.AreEqual(131072, ws.Cell("A29").Value);
            Assert.AreEqual(786432, ws.Cell("A30").Value);
            Assert.AreEqual(1097728, ws.Cell("A33").Value);
            Assert.AreEqual(16834654, ws.Cell("A34").Value);
            Assert.AreEqual(1835008, ws.Cell("A36").Value);
            Assert.AreEqual(6291456, ws.Cell("A39").Value);
            Assert.AreEqual(31457280, ws.Cell("A41").Value);
            Assert.AreEqual(47185920, ws.Cell("A42").Value);
        }

        [Test]
        public void Sum()
        {
            IXLCell cell = new XLWorkbook().AddWorksheet("Sheet1").FirstCell();
            IXLCell fCell = cell.SetValue(1).CellBelow().SetValue(2).CellBelow();
            fCell.FormulaA1 = "sum(A1:A2)";

            Assert.AreEqual(3.0, fCell.Value);
        }

        [Test]
        public void SumDateTimeAndNumber()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell("A1").Value = 1;
                ws.Cell("A2").Value = new DateTime(2018, 1, 1);
                Assert.AreEqual(43102, ws.Evaluate("SUM(A1:A2)"));

                ws.Cell("A1").Value = 2;
                ws.Cell("A2").FormulaA1 = "DATE(2018,1,1)";
                Assert.AreEqual(43103, ws.Evaluate("SUM(A1:A2)"));
            }
        }

        [Test]
        public void SumSq()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"SumSq(3,4)");
            Assert.AreEqual(25.0, actual);
        }

        [Test]
        public void TextConcat()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 1;
            ws.Cell("B1").Value = 1;
            ws.Cell("B2").Value = 1;

            ws.Cell("C1").FormulaA1 = "\"The total value is: \" & SUM(A1:B2)";

            object r = ws.Cell("C1").Value;
            Assert.AreEqual("The total value is: 4", r.ToString());
        }

        [Test]
        public void Trim()
        {
            Assert.AreEqual("Test", XLWorkbook.EvaluateExpr("Trim(\"Test    \")"));

            //Should not trim non breaking space
            //See http://office.microsoft.com/en-us/excel-help/trim-function-HP010062581.aspx
            Assert.AreEqual("Test\u00A0", XLWorkbook.EvaluateExpr("Trim(\"Test\u00A0 \")"));
        }

        [Test]
        public void TestEmptyTallyOperations()
        {
            //In these test no values have been set
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add("TallyTests");
            var cell = wb.Worksheet(1).Cell(1, 1).SetFormulaA1("=MAX(D1,D2)");
            Assert.AreEqual(0, cell.Value);
            cell = wb.Worksheet(1).Cell(2, 1).SetFormulaA1("=MIN(D1,D2)");
            Assert.AreEqual(0, cell.Value);
            cell = wb.Worksheet(1).Cell(3, 1).SetFormulaA1("=SUM(D1,D2)");
            Assert.AreEqual(0, cell.Value);
            Assert.That(() => wb.Worksheet(1).Cell(3, 1).SetFormulaA1("=AVERAGE(D1,D2)").Value, Throws.TypeOf<ApplicationException>());
        }

        [Test]
        public void TestOmittedParameters()
        {
            using (var wb = new XLWorkbook())
            {
                object value;
                value = wb.Evaluate("=IF(TRUE,1)");
                Assert.AreEqual(1, value);

                value = wb.Evaluate("=IF(TRUE,1,)");
                Assert.AreEqual(1, value);

                value = wb.Evaluate("=IF(FALSE,1,)");
                Assert.AreEqual(false, value);

                value = wb.Evaluate("=IF(FALSE,,2)");
                Assert.AreEqual(2, value);
            }
        }

        [Test]
        public void TestDefaultExcelFunctionNamespace()
        {
            Assert.DoesNotThrow(() => XLWorkbook.EvaluateExpr("TODAY()"));
            Assert.DoesNotThrow(() => XLWorkbook.EvaluateExpr("_xlfn.TODAY()"));
            Assert.IsTrue((bool)XLWorkbook.EvaluateExpr("_xlfn.TODAY() = TODAY()"));
        }

        [TestCase("=1234%", 12.34)]
        [TestCase("=1234%%", 0.1234)]
        [TestCase("=100+200%", 102.0)]
        [TestCase("=100%+200", 201.0)]
        [TestCase("=(100+200)%", 3.0)]
        [TestCase("=200%^5", 32.0)]
        [TestCase("=200%^400%", 16.0)]
        [TestCase("=SUM(100,200,300)%", 6.0)]
        public void PercentOperator(string formula, double expectedResult)
        {
            var res = (double)XLWorkbook.EvaluateExpr(formula);

            Assert.AreEqual(expectedResult, res, XLHelper.Epsilon);
        }

        [TestCase("=--1", 1)]
        [TestCase("=++1", 1)]
        [TestCase("=-+-+-1", -1)]
        [TestCase("=2^---2", 0.25)]
        public void MultipleUnaryOperators(string formula, double expectedResult)
        {
            var res = (double)XLWorkbook.EvaluateExpr(formula);

            Assert.AreEqual(expectedResult, res, XLHelper.Epsilon);
        }

        [TestCase("RIGHT(\"2020\", 2) + 1", 21)]
        [TestCase("LEFT(\"20.2020\", 6) + 1", 21.202)]
        [TestCase("2 + (\"3\" & \"4\")", 36)]
        [TestCase("2 + \"3\" & \"4\"", "54")]
        [TestCase("\"7\" & \"4\"", "74")]
        public void TestStringSubExpression(string formula, object expectedResult)
        {
            var actual = XLWorkbook.EvaluateExpr(formula);

            Assert.AreEqual(expectedResult, actual);
        }
    }
}
