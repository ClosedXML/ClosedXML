using ClosedXML.Excel;
using NUnit.Framework;
using System;

namespace ClosedXML.Tests.Excel.CalcEngine
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
            object actual;

            actual = XLWorkbook.EvaluateExpr(@"Asc(""Text"")");
            Assert.That(actual, Is.EqualTo("Text"));
        }

        [Test]
        public void Clean()
        {
            object actual;

            actual = XLWorkbook.EvaluateExpr(string.Format(@"Clean(""A{0}B"")", Environment.NewLine));
            Assert.That(actual, Is.EqualTo("AB"));
        }

        [Test]
        public void Dollar()
        {
            using var wb = new XLWorkbook();
            object actual = wb.Evaluate("DOLLAR(12345.123)");
            Assert.That(actual, Is.EqualTo(TestHelper.CurrencySymbol + "12,345.12"));

            actual = wb.Evaluate("DOLLAR(12345.123, 1)");
            Assert.That(actual, Is.EqualTo(TestHelper.CurrencySymbol + "12,345.1"));
        }

        [TestCase("A", "A", true)]
        [TestCase("A", "a", false)]
        [TestCase("", "", true)]
        public void Exact(string lhs, string rhs, bool result)
        {
            var actual = XLWorkbook.EvaluateExpr($"EXACT(\"{lhs}\", \"{rhs}\")");
            Assert.That(actual, Is.EqualTo(result));
        }

        [Test]
        public void Exact_converts_values_to_text()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr("EXACT(TRUE, \"true\")"), Is.EqualTo(false));
                Assert.That(XLWorkbook.EvaluateExpr("EXACT(TRUE, \"TRUE\")"), Is.EqualTo(true));
                Assert.That(XLWorkbook.EvaluateExpr("EXACT(1, \"1\")"), Is.EqualTo(true));
                Assert.That(XLWorkbook.EvaluateExpr("EXACT(IF(TRUE,), \"\")"), Is.EqualTo(true));
            });

            // Check blank cell
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.That(ws.Evaluate("EXACT(A1, \"\")"), Is.EqualTo(true));
        }

        [Test]
        public void Exact_propagates_errors()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr("EXACT(#DIV/0!, \"A\")"), Is.EqualTo(XLError.DivisionByZero));
                Assert.That(XLWorkbook.EvaluateExpr("EXACT(\"A\", #DIV/0!)"), Is.EqualTo(XLError.DivisionByZero));
            });
        }

        [Test]
        public void Fixed()
        {
            object actual;

            actual = XLWorkbook.EvaluateExpr("Fixed(12345.123)");
            Assert.That(actual, Is.EqualTo("12,345.12"));

            actual = XLWorkbook.EvaluateExpr("Fixed(12345.123, 1)");
            Assert.That(actual, Is.EqualTo("12,345.1"));

            actual = XLWorkbook.EvaluateExpr("Fixed(12345.123, 1, TRUE)");
            Assert.That(actual, Is.EqualTo("12345.1"));
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
            Assert.That(v, Is.EqualTo(3.0));
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
            Assert.That(r, Is.EqualTo("The total value is: 4"));
        }

        [Test]
        public void Trim()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr("Trim(\"Test    \")"), Is.EqualTo("Test"));

                //Should not trim non breaking space
                //See http://office.microsoft.com/en-us/excel-help/trim-function-HP010062581.aspx
                Assert.That(XLWorkbook.EvaluateExpr("Trim(\"Test\u00A0 \")"), Is.EqualTo("Test\u00A0"));
            });
        }

        [Test]
        public void TestEmptyTallyOperations()
        {
            //In these test no values have been set
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add("TallyTests");
            var cell = wb.Worksheet(1).Cell(1, 1).SetFormulaA1("=MAX(D1,D2)");
            Assert.That(cell.Value, Is.EqualTo(0));
            cell = wb.Worksheet(1).Cell(2, 1).SetFormulaA1("=MIN(D1,D2)");
            Assert.That(cell.Value, Is.EqualTo(0));
            cell = wb.Worksheet(1).Cell(3, 1).SetFormulaA1("=SUM(D1,D2)");
            Assert.That(cell.Value, Is.EqualTo(0));
        }

        [Test]
        public void TestOmittedParameters()
        {
            using var wb = new XLWorkbook();
            object value;
            value = wb.Evaluate("=IF(TRUE,1)");
            Assert.That(value, Is.EqualTo(1));

            value = wb.Evaluate("=IF(TRUE,1,)");
            Assert.That(value, Is.EqualTo(1));

            value = wb.Evaluate("=ISBLANK(IF(FALSE,1,))");
            Assert.That(value, Is.EqualTo(true));

            value = wb.Evaluate("=IF(FALSE,,2)");
            Assert.That(value, Is.EqualTo(2));
        }

        [Test]
        public void TestDefaultExcelFunctionNamespace()
        {
            Assert.DoesNotThrow(() => XLWorkbook.EvaluateExpr("TODAY()"));
            Assert.DoesNotThrow(() => XLWorkbook.EvaluateExpr("_xlfn.TODAY()"));
            Assert.That((bool)XLWorkbook.EvaluateExpr("_xlfn.TODAY() = TODAY()"), Is.True);
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

            Assert.That(res, Is.EqualTo(expectedResult).Within(XLHelper.Epsilon));
        }

        [TestCase("=--1", 1)]
        [TestCase("=++1", 1)]
        [TestCase("=-+-+-1", -1)]
        [TestCase("=2^---2", 0.25)]
        public void MultipleUnaryOperators(string formula, double expectedResult)
        {
            var res = (double)XLWorkbook.EvaluateExpr(formula);

            Assert.That(res, Is.EqualTo(expectedResult).Within(XLHelper.Epsilon));
        }

        [TestCase("RIGHT(\"2020\", 2) + 1", 21)]
        [TestCase("LEFT(\"20.2020\", 6) + 1", 21.202)]
        [TestCase("2 + (\"3\" & \"4\")", 36)]
        [TestCase("2 + \"3\" & \"4\"", "54")]
        [TestCase("\"7\" & \"4\"", "74")]
        public void TestStringSubExpression(string formula, object expectedResult)
        {
            var actual = XLWorkbook.EvaluateExpr(formula);

            Assert.That(actual, Is.EqualTo(expectedResult));
        }

        [Test]
        public void Cell_function_is_evaluated_to_reference_error()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").FormulaA1 = "$B$4(5)";

            Assert.That(ws.Cell("A1").Value, Is.EqualTo(XLError.CellReference));
        }
    }
}
