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
        public void Dollar()
        {
            object actual = XLWorkbook.EvaluateExpr("Dollar(12345.123)");
            Assert.AreEqual(TestHelper.CurrencySymbol + "12,345.12", actual);

            actual = XLWorkbook.EvaluateExpr("Dollar(12345.123, 1)");
            Assert.AreEqual(TestHelper.CurrencySymbol + "12,345.1", actual);
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
