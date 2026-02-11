using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Threading;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    [SetCulture("en-US")]
    public class FunctionsTests
    {
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
            using var wb = new XLWorkbook();
            object actual = wb.Evaluate("DOLLAR(12345.123)");
            Assert.AreEqual(TestHelper.CurrencySymbol + "12,345.12", actual);

            actual = wb.Evaluate("DOLLAR(12345.123, 1)");
            Assert.AreEqual(TestHelper.CurrencySymbol + "12,345.1", actual);
        }

        [TestCase("A", "A", true)]
        [TestCase("A", "a", false)]
        [TestCase("", "", true)]
        public void Exact(string lhs, string rhs, bool result)
        {
            var actual = XLWorkbook.EvaluateExpr($"EXACT(\"{lhs}\", \"{rhs}\")");
            Assert.AreEqual(result, actual);
        }

        [Test]
        public void Exact_converts_values_to_text()
        {
            Assert.AreEqual(false, XLWorkbook.EvaluateExpr("EXACT(TRUE, \"true\")"));
            Assert.AreEqual(true, XLWorkbook.EvaluateExpr("EXACT(TRUE, \"TRUE\")"));
            Assert.AreEqual(true, XLWorkbook.EvaluateExpr("EXACT(1, \"1\")"));
            Assert.AreEqual(true, XLWorkbook.EvaluateExpr("EXACT(IF(TRUE,), \"\")"));

            // Check blank cell
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.AreEqual(true, ws.Evaluate("EXACT(A1, \"\")"));
        }

        [Test]
        public void Exact_propagates_errors()
        {
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("EXACT(#DIV/0!, \"A\")"));
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("EXACT(\"A\", #DIV/0!)"));
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
            Assert.AreEqual("The total value is: 4", r);
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

                value = wb.Evaluate("=ISBLANK(IF(FALSE,1,))");
                Assert.AreEqual(true, value);

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

        [Test]
        public void Cell_function_is_evaluated_to_reference_error()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").FormulaA1 = "$B$4(5)";

            Assert.AreEqual(XLError.CellReference, ws.Cell("A1").Value);
        }

        [TestCase("BASE(1E15,30)", "MathTrig.Base")]
        [TestCase("COMBIN(1000,500)", "MathTrig.Combin")]
        [TestCase("COMBINA(100,500)", "MathTrig.CombinA")]
        [TestCase("DECIMAL(\"ZZZ\",26)", "MathTrig.Decimal")]
        [TestCase("GCD(123456,54124)", "MathTrig.Gcd")]
        [TestCase("LCM(123456,54124)", "MathTrig.Lcm")]
        [TestCase("MDETERM({1,2;3,4})", "MathTrig.MDeterm")]
        [TestCase("MINVERSE({1,2;3,4})", "MathTrig.MInverse")]
        [TestCase("MMULT({1},{2})", "MathTrig.MMult")]
        [TestCase("MULTINOMIAL(2,3,4)", "MathTrig.Multinomial")]
        [TestCase("PRODUCT(2,3,{4,5,6})", "MathTrig.Product")]
        [TestCase("SERIESSUM(2,1,2,{1,2,3})", "MathTrig.SeriesSum")]
        [TestCase("SUM(2,1,2,{1,2,3})", "MathTrig.Sum")]
        [TestCase("SUMIF(B1:B4,\"=5\")", "MathTrig.SumIf")]
        [TestCase("SUMIFS(B1:B4,C1:C4,\">0\")", "MathTrig.SumIfs")]
        [TestCase("SUMPRODUCT({2,3},{4,5})", "MathTrig.SumProduct")]
        [TestCase("SUMSQ(5,4)", "MathTrig.SumSq")]
        [TestCase("NETWORKDAYS(10,100,{20,50})", "DateAndTime.NetWorkDays")]
        [TestCase("WORKDAY(10,100,{20,50})", "DateAndTime.Workday")]
        [TestCase("YEARFRAC(1,10000,1)", "DateAndTime.YearFrac")]
        [TestCase("AND(TRUE,TRUE)", "Logical.And")]
        [TestCase("OR(TRUE,TRUE)", "Logical.Or")]
        [TestCase("COLUMN(D:Z)", "Lookup.Column")]
        [TestCase("HLOOKUP(2,{0,1,2,3},1)", "Lookup.Hlookup")]
        [TestCase("MATCH(5,{1,7})", "Lookup.Match")]
        [TestCase("ROW(2:5)", "Lookup.Row")]
        [TestCase("VLOOKUP(2,{0;1;2;3},1)", "Lookup.Vlookup")]
        [TestCase("AVERAGE({1,2;3,4})", "Statistical.Average")]
        [TestCase("AVERAGEA({1,2;3,4})", "Statistical.AverageA")]
        [TestCase("BINOMDIST(6,10,0.5,TRUE)", "Statistical.BinomDist")]
        [TestCase("BINOM.DIST(6,10,0.5,TRUE)", "Statistical.BinomDist")]
        [TestCase("COUNT({1,2})", "Statistical.Count")]
        [TestCase("COUNTA({1,2})", "Statistical.Count")]
        [TestCase("COUNTBLANK(D1:D7)", "Statistical.CountBlank")]
        [TestCase("COUNTIF(D1:D7,\">0\")", "Statistical.CountIf")]
        [TestCase("COUNTIFS(D1:D7,\">0\")", "Statistical.CountIfs")]
        [TestCase("DEVSQ({1,2})", "Statistical.DevSq")]
        [TestCase("GEOMEAN({1,2})", "Statistical.GeoMean")]
        [TestCase("LARGE({10,11,12},2)", "Statistical.Large")]
        [TestCase("MAX({1,5})", "Statistical.Max")]
        [TestCase("MAXA({1,5})", "Statistical.MaxA")]
        [TestCase("MEDIAN({1,5,7})", "Statistical.Median")]
        [TestCase("MIN({1,5})", "Statistical.Min")]
        [TestCase("MINA({1,5})", "Statistical.MinA")]
        [TestCase("STDEV({1,5})", "Statistical.StDev")]
        [TestCase("STDEVA({1,5})", "Statistical.StDevA")]
        [TestCase("STDEVP({1,5})", "Statistical.StDevP")]
        [TestCase("STDEVPA({1,5})", "Statistical.StDevPA")]
        [TestCase("STDEV.S({1,5})", "Statistical.StDev")]
        [TestCase("STDEV.P({1,5})", "Statistical.StDevP")]
        [TestCase("VAR({1,5})", "Statistical.Var")]
        [TestCase("VARA({1,5})", "Statistical.VarA")]
        [TestCase("VARP({1,5})", "Statistical.VarP")]
        [TestCase("VARPA({1,5})", "Statistical.VarPA")]
        [TestCase("VAR.S({1,5})", "Statistical.Var")]
        [TestCase("VAR.P({1,5})", "Statistical.VarP")]
        [TestCase("ASC(\"A\")", "Text.Asc")]
        [TestCase("CLEAN(\"A\")", "Text.Clean")]
        [TestCase("CONCAT(\"A\",\"B\")", "Text.Concat")]
        [TestCase("CONCATENATE(\"A\",\"B\")", "Text.Concatenate")]
        [TestCase("LEFT(\"AB\",1)", "Text.Left")]
        [TestCase("LOWER(\"A\")", "Text.Lower")]
        [TestCase("NUMBERVALUE(\"1\")", "Text.NumberValue")]
        [TestCase("PROPER(\"hello\")", "Text.Proper")]
        [TestCase("REPT(\"A\",10)", "Text.Rept")]
        [TestCase("RIGHT(\"AB\",1)", "Text.Right")]
        [TestCase("T({100})", "Text.T")]
        [TestCase("TEXTJOIN(\"-\",TRUE,\"A\",\"B\")", "Text.TextJoin")]
        [TestCase("TRIM(\"ABC\")", "Text.Trim")]
        public void Can_cancel_function_execution(string formula, string expectedStackTrace)
        {
            var cts = new CancellationTokenSource();
            using var wb = new XLWorkbook(new LoadOptions { CancellationToken = cts.Token });
            var ws = wb.AddWorksheet();
            ws.Cell("D1").Value = 4; // Need a non-blank cell for COUNTBLANK check
            ws.Cell("A1").FormulaA1 = formula;

            cts.Cancel();
            var ex = Assert.Throws<OperationCanceledException>(() => _ = ws.Cell("A1").Value);

            StringAssert.Contains(expectedStackTrace + "(", ex?.StackTrace);
        }
    }
}
