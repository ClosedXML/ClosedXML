// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;
using System;
using System.Linq;
using ClosedXML.Excel.CalcEngine.Exceptions;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class StatisticalTests
    {
        private double tolerance = 1e-6;
        private XLWorkbook workbook;

        [Test]
        public void Average()
        {
            double value;
            value = (double)workbook.Evaluate("AVERAGE(-27.5,93.93,64.51,-70.56)");
            Assert.AreEqual(15.095, value, tolerance);

            var ws = workbook.Worksheets.First();
            value = (double)ws.Evaluate("AVERAGE(G3:G45)");
            Assert.AreEqual(49.3255814, value, tolerance);

            Assert.That(() => ws.Evaluate("AVERAGE(D3:D45)"), Throws.TypeOf<ApplicationException>());
        }

        [TestCase(6, 10, 0.5, 0.205078125)]
        [TestCase(4, 20, 0.2, 0.2181994)] // p different than 0.5
        [TestCase(0, 5, 0.2, 0.32768)] // 0 out of 5 successes
        [TestCase(0, 0, 0.2, 1)] // 0 out of 0 successes
        [TestCase(1, 1, 0, 0)]
        [TestCase(1, 1, 1, 1)]
        [TestCase(2, 4, 0.5, 0.375)]
        [TestCase(2.9, 4.9, 0.5, 0.375)] // Attempts are floored
        public void BinomDist_calculates_non_cumulative_binomial_distribution(double k, double n, double p, double expected)
        {
            var kString = k.ToInvariantString();
            var nString = n.ToInvariantString();
            var pString = p.ToInvariantString();
            var result = (double)XLWorkbook.EvaluateExpr($"BINOMDIST({kString}, {nString}, {pString}, FALSE)");
            Assert.AreEqual(expected, result, tolerance);
        }

        [TestCase(6, 10, 0.5, 0.828125)]
        [TestCase(2, 7, 0.3, 0.6470695)]
        [TestCase(0, 7, 0.3, 0.0823543)]
        [TestCase(0, 0, 0.3, 1)]
        [TestCase(0, 0, 1, 1)]
        [TestCase(2, 4, 0.5, 0.6875)]
        [TestCase(2.9, 4.9, 0.5, 0.6875)] // Values are floored
        public void BinomDist_calculates_cumulative_binomial_distribution(double k, double n, double p, double expected)
        {
            var kString = k.ToInvariantString();
            var nString = n.ToInvariantString();
            var pString = p.ToInvariantString();
            var result = (double)XLWorkbook.EvaluateExpr($"BINOMDIST({kString}, {nString}, {pString}, TRUE)");
            Assert.AreEqual(expected, result, tolerance);
        }

        [TestCase(5, 4, 0.5)] // Five successes out of 4 attempts
        [TestCase(-1, 4, 0.5)] // Negative successes
        [TestCase(0, -1, 0.5)] // Negative attempts
        [TestCase(2, 4, -0.1)] // p < 0
        [TestCase(2, 4, 1.1)] // p > 1
        [TestCase(1E+300, 2E+300, 0.5)] // Too large values
        public void BinomDist_returns_num_error_on_invalid_calculations(double k, double n, double p)
        {
            var kString = k.ToInvariantString();
            var nString = n.ToInvariantString();
            var pString = p.ToInvariantString();
            var result = XLWorkbook.EvaluateExpr($"BINOMDIST({kString}, {nString}, {pString}, FALSE)");
            Assert.AreEqual(XLError.NumberInvalid, result);
        }

        [Test]
        public void Count()
        {
            var ws = workbook.Worksheets.First();
            XLCellValue value;
            value = ws.Evaluate(@"=COUNT(D3:D45)");
            Assert.AreEqual(0, value);

            value = ws.Evaluate(@"=COUNT(G3:G45)");
            Assert.AreEqual(43, value);

            value = ws.Evaluate(@"=COUNT(G:G)");
            Assert.AreEqual(43, value);

            value = workbook.Evaluate(@"=COUNT(Data!G:G)");
            Assert.AreEqual(43, value);
        }

        [Test]
        public void CountA()
        {
            var ws = workbook.Worksheets.First();
            var value = ws.Evaluate("COUNTA(D3:D45)");
            Assert.AreEqual(43, value);

            value = ws.Evaluate("COUNTA(G3:G45)");
            Assert.AreEqual(43, value);

            value = ws.Evaluate("COUNTA(G:G)");
            Assert.AreEqual(44, value);

            value = workbook.Evaluate("COUNTA(Data!G:G)");
            Assert.AreEqual(44, value);
        }

        [Test]
        public void CountA_counts_non_blank_values()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = Blank.Value;
            ws.Cell("A2").Value = 39790;
            ws.Cell("A3").Value = 0;
            ws.Cell("A4").Value = 22.24;
            ws.Cell("A5").Value = "Text";
            ws.Cell("A6").Value = false;
            ws.Cell("A7").Value = true;
            ws.Cell("A8").Value = XLError.DivisionByZero;
            ws.Cell("A9").FormulaA1 = "COUNTA(A1:B8)";
            Assert.AreEqual(7, ws.Cell("A9").Value);
        }

        [Test]
        public void CountA_on_examples_from_spec()
        {
            Assert.AreEqual(5, XLWorkbook.EvaluateExpr("COUNTA(1,2,3,4,5)"));
            Assert.AreEqual(5, XLWorkbook.EvaluateExpr("COUNTA(1,2,3,4,5)"));
            Assert.AreEqual(7, XLWorkbook.EvaluateExpr("COUNTA({1,2,3,4,5},6,\"7\")"));

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("E2").Value = true;
            Assert.AreEqual(1, ws.Evaluate("COUNTA(10, E1)"));
            Assert.AreEqual(2, ws.Evaluate("COUNTA(10, E2)"));
        }

        [Test]
        public void CountA_accepts_union_references()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A2").Value = 7;
            ws.Cell("B5").Value = false;
            Assert.AreEqual(2, ws.Evaluate("COUNTA((A1:A4,B4:B7))"));
        }

        [Test]
        public void CountA_doesnt_count_single_blank_cell_reference()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.AreEqual(0, ws.Evaluate("COUNTA(A1)"));
        }

        [Test]
        public void CountA_counts_blank_argument()
        {
            Assert.AreEqual(1, XLWorkbook.EvaluateExpr("COUNTA(IF(TRUE,,))"));
        }

        [Test]
        public void CountA_counts_error_arguments()
        {
            Assert.AreEqual(7, XLWorkbook.EvaluateExpr("COUNTA(#NULL!, #DIV/0!, #VALUE!, #REF!, #NAME?, #NUM!, #N/A)"));
        }

        [Test]
        public void CountA_counts_empty_string()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = string.Empty;
            Assert.AreEqual(2, ws.Evaluate("COUNTA(A1, \"\")"));
        }

        [Test]
        public void CountBlank()
        {
            var ws = workbook.Worksheets.First();
            XLCellValue value;
            value = ws.Evaluate(@"=COUNTBLANK(B:B)");
            Assert.AreEqual(1048532, value);

            value = ws.Evaluate(@"=COUNTBLANK(D43:D49)");
            Assert.AreEqual(4, value);

            value = ws.Evaluate(@"=COUNTBLANK(E3:E45)");
            Assert.AreEqual(0, value);

            value = ws.Evaluate(@"=COUNTBLANK(A1)");
            Assert.AreEqual(1, value);

            Assert.Throws<MissingContextException>(() => workbook.Evaluate(@"=COUNTBLANK(E3:E45)"));
            Assert.Throws<ExpressionParseException>(() => ws.Evaluate(@"=COUNTBLANK()"));
            Assert.Throws<ExpressionParseException>(() => ws.Evaluate(@"=COUNTBLANK(A3:A45,E3:E45)"));
        }

        [Test]
        public void CountIf()
        {
            var ws = workbook.Worksheets.First();
            XLCellValue value;
            value = ws.Evaluate(@"=COUNTIF(D3:D45,""Central"")");
            Assert.AreEqual(24, value);

            value = ws.Evaluate(@"=COUNTIF(D:D,""Central"")");
            Assert.AreEqual(24, value);

            value = workbook.Evaluate(@"=COUNTIF(Data!D:D,""Central"")");
            Assert.AreEqual(24, value);
        }

        [TestCase(@"=COUNTIF(Data!E:E, ""J*"")", 13)]
        [TestCase(@"=COUNTIF(Data!E:E, ""*i*"")", 21)]
        [TestCase(@"=COUNTIF(Data!E:E, ""*in*"")", 9)]
        [TestCase(@"=COUNTIF(Data!E:E, ""*i*l"")", 9)]
        [TestCase(@"=COUNTIF(Data!E:E, ""*i?e*"")", 9)]
        [TestCase(@"=COUNTIF(Data!E:E, ""*o??s*"")", 10)]
        [TestCase(@"=COUNTIF(Data!X1:X1000, """")", 1000)]
        [TestCase(@"=COUNTIF(Data!E1:E44, """")", 1)]
        public void CountIf_ConditionWithWildcards(string formula, int expectedResult)
        {
            var ws = workbook.Worksheets.First();

            var value = ws.Evaluate(formula);
            Assert.AreEqual(expectedResult, value);
        }

        [TestCase(@"=COUNTIF(A1:A10, 1)", 1)]
        [TestCase(@"=COUNTIF(A1:A10, 2.0)", 1)]
        [TestCase(@"=COUNTIF(A1:A10, ""3"")", 2)]
        [TestCase(@"=COUNTIF(A1:A10, 3)", 2)]
        [TestCase(@"=COUNTIF(A1:A10, 43831)", 1)]
        [TestCase(@"=COUNTIF(A1:A10, DATE(2020, 1, 1))", 1)]
        [TestCase(@"=COUNTIF(A1:A10, TRUE)", 1)]
        public void CountIf_MixedData(string formula, int expected)
        {
            // We follow to Excel's convention.
            // Excel treats 1 and TRUE as unequal, but 3 and "3" as equal
            // LibreOffice Calc handles some SUMIF and COUNTIF differently, e.g. it treats 1 and TRUE as equal, but 3 and "3" differently
            var ws = workbook.Worksheet("MixedData");
            Assert.AreEqual(expected, ws.Evaluate(formula));
        }

        [TestCase("x", @"=COUNTIF(A1:A1, ""?"")", 1)]
        [TestCase("x", @"=COUNTIF(A1:A1, ""~?"")", 0)]
        [TestCase("?", @"=COUNTIF(A1:A1, ""~?"")", 1)]
        [TestCase("~?", @"=COUNTIF(A1:A1, ""~?"")", 0)]
        [TestCase("~?", @"=COUNTIF(A1:A1, ""~~~?"")", 1)]
        [TestCase("?", @"=COUNTIF(A1:A1, ""~~?"")", 0)]
        [TestCase("~?", @"=COUNTIF(A1:A1, ""~~?"")", 1)]
        [TestCase("~x", @"=COUNTIF(A1:A1, ""~~?"")", 1)]
        [TestCase("*", @"=COUNTIF(A1:A1, ""~*"")", 1)]
        [TestCase("~*", @"=COUNTIF(A1:A1, ""~*"")", 0)]
        [TestCase("~*", @"=COUNTIF(A1:A1, ""~~~*"")", 1)]
        [TestCase("*", @"=COUNTIF(A1:A1, ""~~*"")", 0)]
        [TestCase("~*", @"=COUNTIF(A1:A1, ""~~*"")", 1)]
        [TestCase("~x", @"=COUNTIF(A1:A1, ""~~*"")", 1)]
        [TestCase("~xyz", @"=COUNTIF(A1:A1, ""~~*"")", 1)]
        public void CountIf_MoreWildcards(string cellContent, string formula, int expectedResult)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                ws.Cell(1, 1).Value = cellContent;

                Assert.AreEqual(expectedResult, (double)ws.Evaluate(formula));
            }
        }

        [TestCase("=COUNTIFS(B1:D1, \"=Yes\")", 1)]
        [TestCase("=COUNTIFS(B1:B4, \"=Yes\", C1:C4, \"=Yes\")", 2)]
        [TestCase("= COUNTIFS(B4:D4, \"=Yes\", B2:D2, \"=Yes\")", 1)]
        public void CountIfs_ReferenceExample1FromExcelDocumentations(
            string formula,
            int expectedOutcome)
        {
            using (var wb = new XLWorkbook())
            {
                wb.ReferenceStyle = XLReferenceStyle.A1;

                var ws = wb.AddWorksheet("Sheet1");

                ws.Cell(1, 1).Value = "Davidoski";
                ws.Cell(1, 2).Value = "Yes";
                ws.Cell(1, 3).Value = "No";
                ws.Cell(1, 4).Value = "No";

                ws.Cell(2, 1).Value = "Burke";
                ws.Cell(2, 2).Value = "Yes";
                ws.Cell(2, 3).Value = "Yes";
                ws.Cell(2, 4).Value = "No";

                ws.Cell(3, 1).Value = "Sundaram";
                ws.Cell(3, 2).Value = "Yes";
                ws.Cell(3, 3).Value = "Yes";
                ws.Cell(3, 4).Value = "Yes";

                ws.Cell(4, 1).Value = "Levitan";
                ws.Cell(4, 2).Value = "No";
                ws.Cell(4, 3).Value = "Yes";
                ws.Cell(4, 4).Value = "Yes";

                Assert.AreEqual(expectedOutcome, ws.Evaluate(formula));
            }
        }

        [Test]
        public void CountIfs_SingleCondition()
        {
            var ws = workbook.Worksheets.First();
            XLCellValue value;
            value = ws.Evaluate(@"=COUNTIFS(D3:D45,""Central"")");
            Assert.AreEqual(24, value);

            value = ws.Evaluate(@"=COUNTIFS(D:D,""Central"")");
            Assert.AreEqual(24, value);

            value = workbook.Evaluate(@"=COUNTIFS(Data!D:D,""Central"")");
            Assert.AreEqual(24, value);
        }

        [TestCase(@"=COUNTIFS(Data!E:E, ""J*"")", 13)]
        [TestCase(@"=COUNTIFS(Data!E:E, ""*i*"")", 21)]
        [TestCase(@"=COUNTIFS(Data!E:E, ""*in*"")", 9)]
        [TestCase(@"=COUNTIFS(Data!E:E, ""*i*l"")", 9)]
        [TestCase(@"=COUNTIFS(Data!E:E, ""*i?e*"")", 9)]
        [TestCase(@"=COUNTIFS(Data!E:E, ""*o??s*"")", 10)]
        [TestCase(@"=COUNTIFS(Data!X1:X1000, """")", 1000)]
        [TestCase(@"=COUNTIFS(Data!E1:E44, """")", 1)]
        public void CountIfs_SingleConditionWithWildcards(string formula, int expectedResult)
        {
            var ws = workbook.Worksheets.First();

            var value = ws.Evaluate(formula);
            Assert.AreEqual(expectedResult, value);
        }

        [OneTimeTearDown]
        public void Dispose()
        {
            workbook.Dispose();
        }

        [TestCase(@"H3:H45", ExpectedResult = 7.51126069234216)]
        [TestCase(@"H:H", ExpectedResult = 7.51126069234216)]
        [TestCase(@"Data!H:H", ExpectedResult = 7.51126069234216)]
        [TestCase(@"H3:H10", ExpectedResult = 5.26214814727941)]
        [TestCase(@"H3:H20", ExpectedResult = 7.01281435054797)]
        [TestCase(@"H3:H30", ExpectedResult = 7.00137389296182)]
        [TestCase(@"H3:H3", ExpectedResult = 1.99)]
        [TestCase(@"H10:H20", ExpectedResult = 8.37855107505682)]
        [TestCase(@"H15:H20", ExpectedResult = 15.8927310267677)]
        [TestCase(@"H20:H30", ExpectedResult = 7.14321227391814)]
        [DefaultFloatingPointTolerance(1e-12)]
        public double Geomean(string sourceValue)
        {
            return (double)workbook.Worksheets.First().Evaluate($"=GEOMEAN({sourceValue})");
        }

        [TestCase("D3:D45", ExpectedResult = XLError.NumberInvalid)]
        [TestCase("-1, 0, 3", ExpectedResult = XLError.NumberInvalid)]
        public XLError Geomean_IncorrectCases(string sourceValue)
        {
            var ws = workbook.Worksheets.First();

            return (XLError)ws.Evaluate($"GEOMEAN({sourceValue})");
        }

        [SetUp]
        public void Init()
        {
            // Make sure tests run on a deterministic culture
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            workbook = SetupWorkbook();
        }

        [TestCase(@"H3:H45", ExpectedResult = 94145.5271162791)]
        [TestCase(@"H:H", ExpectedResult = 94145.5271162791)]
        [TestCase(@"Data!H:H", ExpectedResult = 94145.5271162791)]
        [TestCase(@"H3:H10", ExpectedResult = 411.5)]
        [TestCase(@"H3:H20", ExpectedResult = 13604.2067611111)]
        [TestCase(@"H3:H30", ExpectedResult = 14231.0694)]
        [TestCase(@"H3:H3", ExpectedResult = 0)]
        [TestCase(@"H10:H20", ExpectedResult = 12713.7600909091)]
        [TestCase(@"H15:H20", ExpectedResult = 10827.2200833333)]
        [TestCase(@"H20:H30", ExpectedResult = 477.132272727273)]
        [DefaultFloatingPointTolerance(1e-10)]
        public double DevSq(string sourceValue)
        {
            return (double)workbook.Worksheets.First().Evaluate($"=DEVSQ({sourceValue})");
        }

        [TestCase("D3:D45", ExpectedResult = XLError.IncompatibleValue)]
        public XLError Devsq_IncorrectCases(string sourceValue)
        {
            var ws = workbook.Worksheets.First();

            return (XLError)ws.Evaluate($"DEVSQ({sourceValue})");
        }

        [TestCase(0, ExpectedResult = 0)]
        [TestCase(0.2, ExpectedResult = 0.202732554054082)]
        [TestCase(0.25, ExpectedResult = 0.255412811882995)]
        [TestCase(0.3296001056, ExpectedResult = 0.342379555936801)]
        [TestCase(-0.36, ExpectedResult = -0.37688590118819)]
        [TestCase(-0.000003, ExpectedResult = -0.00000299999999998981)]
        [TestCase(-0.063453535345348, ExpectedResult = -0.0635389037459617)]
        [TestCase(0.559015883901589171354964, ExpectedResult = 0.631400600322212)]
        [TestCase(0.2691496, ExpectedResult = 0.275946780611959)]
        [TestCase(-0.10674142, ExpectedResult = -0.107149608461448)]
        [DefaultFloatingPointTolerance(1e-12)]
        public double Fisher(double sourceValue)
        {
            return (double)XLWorkbook.EvaluateExpr($"FISHER({sourceValue})");
        }

        // TODO : the string case will be treated correctly when Coercion is implemented better
        //[TestCase("asdf", ExpectedResult = XLError.IncompatibleValue)]
        [TestCase("5", ExpectedResult = XLError.NumberInvalid)]
        [TestCase("-1", ExpectedResult = XLError.NumberInvalid)]
        [TestCase("1", ExpectedResult = XLError.NumberInvalid)]
        public XLError Fisher_IncorrectCases(string sourceValue)
        {
            return (XLError)XLWorkbook.EvaluateExpr($"FISHER({sourceValue})");
        }

        [Test]
        public void Max()
        {
            var ws = workbook.Worksheets.First();
            XLCellValue value;
            value = ws.Evaluate(@"=MAX(D3:D45)");
            Assert.AreEqual(0, value);

            value = ws.Evaluate(@"=MAX(G3:G45)");
            Assert.AreEqual(96, value);

            value = ws.Evaluate(@"=MAX(G:G)");
            Assert.AreEqual(96, value);

            value = workbook.Evaluate(@"=MAX(Data!G:G)");
            Assert.AreEqual(96, value);

            // Although in most cases blank cells are considered 0, MAX just ignores them.
            value = workbook.Evaluate(@"MAX(-10, Data!X:Z)");
            Assert.AreEqual(-10, value);

            // Blanks are not ignored as a value, only in references.
            value = workbook.Evaluate(@"MAX(-10, IF(TRUE,,))");
            Assert.AreEqual(0, value);

            // Logical are converted
            value = workbook.Evaluate(@"MAX(-10, TRUE)");
            Assert.AreEqual(1, value);

            // Numbers texts are converted
            value = workbook.Evaluate(@"MAX(-10, ""10"")");
            Assert.AreEqual(10, value);

            // Non-number texts cause conversion error
            value = workbook.Evaluate(@"MAX(-10, ""a"")");
            Assert.AreEqual(XLError.IncompatibleValue, value);

            // Arrays - numbers are used
            value = workbook.Evaluate(@"MAX(-10, { -6, -5, 7 })");
            Assert.AreEqual(7, value);

            // Arrays - non-number and non-error values are skipped.
            value = workbook.Evaluate(@"MAX(-10, { TRUE, FALSE, ""100"" })");
            Assert.AreEqual(-10, value);

            // Arrays - errors immediately end evaluation.
            value = workbook.Evaluate(@"MAX(-10, {#N/A})");
            Assert.AreEqual(XLError.NoValueAvailable, value);
        }

        [Test]
        public void Median_CellRangeOfNonNumericValues_ThrowsApplicationException()
        {
            //Arrange
            var ws = workbook.Worksheets.First();

            //Act - Assert
            Assert.Throws<ApplicationException>(() =>
            {
                ws.Evaluate("AVERAGE(D3:D45)");
            });
        }

        [Test]
        public void Median_EvenCountOfCellRange_ReturnsAverageOfTwoElementsInMiddleOfSortedList()
        {
            //Arrange
            var ws = workbook.Worksheets.First();

            //Act
            var value = (double)ws.Evaluate("MEDIAN(I3:I10)");

            //Assert
            Assert.AreEqual(244.225, value, tolerance);
        }

        [Test]
        public void Median_EvenCountOfManualNumbers_ReturnsAverageOfTwoElementsInMiddleOfSortedList()
        {
            //Act
            var value = (double)workbook.Evaluate("MEDIAN(-27.5,93.93,64.51,-70.56)");

            //Assert
            Assert.AreEqual(18.505, value, tolerance);
        }

        [Test]
        public void Median_OddCountOfCellRange_ReturnsElementInMiddleOfSortedList()
        {
            //Arrange
            var ws = workbook.Worksheets.First();

            //Act
            var value = (double)ws.Evaluate("MEDIAN(I3:I11)");

            //Assert
            Assert.AreEqual(189.05, value, tolerance);
        }

        [Test]
        public void Median_OddCountOfManualNumbers_ReturnsElementInMiddleOfSortedList()
        {
            //Act
            var value = (double)workbook.Evaluate("MEDIAN(-27.5,93.93,64.51,-70.56,101.65)");

            //Assert
            Assert.AreEqual(64.51, value, tolerance);
        }

        [Test]
        public void Min()
        {
            var ws = workbook.Worksheets.First();
            XLCellValue value;
            value = ws.Evaluate(@"=MIN(D3:D45)");
            Assert.AreEqual(0, value);

            value = ws.Evaluate(@"=MIN(G3:G45)");
            Assert.AreEqual(2, value);

            value = ws.Evaluate(@"=MIN(G:G)");
            Assert.AreEqual(2, value);

            value = workbook.Evaluate(@"=MIN(Data!G:G)");
            Assert.AreEqual(2, value);
        }

        [Test]
        public void StDev()
        {
            var ws = workbook.Worksheets.First();
            double value;
            Assert.That(() => ws.Evaluate(@"=STDEV(D3:D45)"), Throws.TypeOf<ApplicationException>());

            value = (double)ws.Evaluate(@"=STDEV(H3:H45)");
            Assert.AreEqual(47.34511769, value, tolerance);

            value = (double)ws.Evaluate(@"=STDEV(H:H)");
            Assert.AreEqual(47.34511769, value, tolerance);

            value = (double)workbook.Evaluate(@"=STDEV(Data!H:H)");
            Assert.AreEqual(47.34511769, value, tolerance);
        }

        [Test]
        public void StDevP()
        {
            var ws = workbook.Worksheets.First();
            double value;
            Assert.That(() => ws.Evaluate(@"=STDEVP(D3:D45)"), Throws.InvalidOperationException);

            value = (double)ws.Evaluate(@"=STDEVP(H3:H45)");
            Assert.AreEqual(46.79135458, value, tolerance);

            value = (double)ws.Evaluate(@"=STDEVP(H:H)");
            Assert.AreEqual(46.79135458, value, tolerance);

            value = (double)workbook.Evaluate(@"=STDEVP(Data!H:H)");
            Assert.AreEqual(46.79135458, value, tolerance);
        }

        [TestCase(@"=SUMIF(A1:A10, 1, A1:A10)", 1)]
        [TestCase(@"=SUMIF(A1:A10, 2.0, A1:A10)", 2)]
        [TestCase(@"=SUMIF(A1:A10, 3, A1:A10)", 3)]
        [TestCase(@"=SUMIF(A1:A10, ""3"", A1:A10)", 3)]
        [TestCase(@"=SUMIF(A1:A10, 43831, A1:A10)", 43831)]
        [TestCase(@"=SUMIF(A1:A10, DATE(2020, 1, 1), A1:A10)", 43831)]
        [TestCase(@"=SUMIF(A1:A10, TRUE, A1:A10)", 0)]
        public void SumIf_MixedData(string formula, double expected)
        {
            // We follow to Excel's convention.
            // Excel treats 1 and TRUE as unequal, but 3 and "3" as equal
            // LibreOffice Calc handles some SUMIF and COUNTIF differently, e.g. it treats 1 and TRUE as equal, but 3 and "3" differently
            var ws = workbook.Worksheet("MixedData");
            Assert.AreEqual(expected, ws.Evaluate(formula));
        }

        [Test]
        [TestCase("COUNT(G:I,G:G,H:I)", 258d, Description = "COUNT overlapping columns")]
        [TestCase("COUNT(6:8,6:6,7:8)", 30d, Description = "COUNT overlapping rows")]
        [TestCase("COUNTBLANK(H:J)", 3145640d, Description = "COUNTBLANK columns")]
        [TestCase("COUNTBLANK(7:9)", 49128d, Description = "COUNTBLANK rows")]
        [TestCase("COUNT(1:1048576)", 216d, Description = "COUNT worksheet")]
        [TestCase("COUNTBLANK(1:1048576)", 17179868831d, Description = "COUNTBLANK worksheet")]
        [TestCase("SUM(H:J)", 20501.15d, Description = "SUM columns")]
        [TestCase("SUM(4:5)", 85366.12d, Description = "SUM rows")]
        [TestCase("SUMIF(G:G,50,H:H)", 24.98d, Description = "SUMIF columns")]
        [TestCase("SUMIF(G23:G52,\"\",H3:H32)", 53.24d, Description = "SUMIF ranges")]
        [TestCase("SUMIFS(H:H,G:G,50,I:I,\">900\")", 19.99d, Description = "SUMIFS columns")]
        public void TallySkipsEmptyCells(string formulaA1, double expectedResult)
        {
            using (var wb = SetupWorkbook())
            {
                var ws = wb.Worksheets.First();
                //Let's pre-initialize cells we need so they didn't affect the result
                ws.Range("A1:J45").Style.Fill.BackgroundColor = XLColor.Amber;
                ws.Cell("ZZ1000").Value = 1;

                var actualResult = (double)ws.Evaluate(formulaA1);

                Assert.AreEqual(expectedResult, actualResult, tolerance);
            }
        }

        [Test]
        public void Var()
        {
            var ws = workbook.Worksheets.First();
            double value;
            Assert.That(() => ws.Evaluate(@"=VAR(D3:D45)"), Throws.InvalidOperationException);

            value = (double)ws.Evaluate(@"=VAR(H3:H45)");
            Assert.AreEqual(2241.560169, value, tolerance);

            value = (double)ws.Evaluate(@"=VAR(H:H)");
            Assert.AreEqual(2241.560169, value, tolerance);

            value = (double)workbook.Evaluate(@"=VAR(Data!H:H)");
            Assert.AreEqual(2241.560169, value, tolerance);
        }

        [Test]
        public void VarP()
        {
            var ws = workbook.Worksheets.First();
            double value;
            Assert.That(() => ws.Evaluate(@"=VARP(D3:D45)"), Throws.InvalidOperationException);

            value = (double)ws.Evaluate(@"=VARP(H3:H45)");
            Assert.AreEqual(2189.430863, value, tolerance);

            value = (double)ws.Evaluate(@"=VARP(H:H)");
            Assert.AreEqual(2189.430863, value, tolerance);

            value = (double)workbook.Evaluate(@"=VARP(Data!H:H)");
            Assert.AreEqual(2189.430863, value, tolerance);
        }

        [Test]
        public void Large()
        {
            var ws = workbook.Worksheet("Data");
            var value = ws.Evaluate("LARGE(G1:G45, 1)");
            Assert.AreEqual(96, value);

            value = ws.Evaluate("LARGE(G1:G45, 7)");
            Assert.AreEqual(87, value);

            value = ws.Evaluate("LARGE(G1:G45, 0)");
            Assert.AreEqual(XLError.NumberInvalid, value);

            value = ws.Evaluate("LARGE(G1:G45, -1)");
            Assert.AreEqual(XLError.NumberInvalid, value);

            value = ws.Evaluate("LARGE(G1:G45,\"test\")");
            Assert.AreEqual(XLError.IncompatibleValue, value);

            value = ws.Evaluate("LARGE(C:C,7)");
            Assert.AreEqual(42623, value);

            value = ws.Evaluate("LARGE(D:D,7)");
            Assert.AreEqual(XLError.NumberInvalid, value);

            ws = workbook.Worksheet("MixedData");

            value = ws.Evaluate("LARGE(A1:A7,6)");
            Assert.AreEqual(XLError.NumberInvalid, value);

            // Ignores non-numbers.
            value = ws.Evaluate("LARGE(A1:A7,5)");
            Assert.AreEqual(1, value);

            // Accepts non-area references.
            value = ws.Evaluate("LARGE((A1:A2,A4:A6),2)");
            Assert.AreEqual(3, value);

            // Errors are returned.
            value = ws.Evaluate("LARGE({ 1, 2, #N/A }, 1)");
            Assert.AreEqual(XLError.NoValueAvailable, value);

            // Uses ceiling logic for number (1.1 -> 2) + can use arrays.
            value = ws.Evaluate("LARGE({ 1, 2 }, 1.1)");
            Assert.AreEqual(1, value);

            // If a scalar number-like value supplied, it is converted to number.
            value = ws.Evaluate("LARGE(\"1 1/2\", 1)");
            Assert.AreEqual(1.5, value);

            // When the scalar can't be converted, return conversion error.
            value = ws.Evaluate("LARGE(\"test\", 1)");
            Assert.AreEqual(XLError.IncompatibleValue, value);
        }

        private XLWorkbook SetupWorkbook()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Data");
            var data = new object[]
            {
                new {Id=1, OrderDate = DateTime.Parse("2015-01-06"), Region = "East", Rep = "Jones", Item = "Pencil", Units = 95, UnitCost = 1.99, Total = 189.05 },
                new {Id=2, OrderDate = DateTime.Parse("2015-01-23"), Region = "Central", Rep = "Kivell", Item = "Binder", Units = 50, UnitCost = 19.99, Total = 999.5},
                new {Id=3, OrderDate = DateTime.Parse("2015-02-09"), Region = "Central", Rep = "Jardine", Item = "Pencil", Units = 36, UnitCost = 4.99, Total = 179.64},
                new {Id=4, OrderDate = DateTime.Parse("2015-02-26"), Region = "Central", Rep = "Gill", Item = "Pen", Units = 27, UnitCost = 19.99, Total = 539.73},
                new {Id=5, OrderDate = DateTime.Parse("2015-03-15"), Region = "West", Rep = "Sorvino", Item = "Pencil", Units = 56, UnitCost = 2.99, Total = 167.44},
                new {Id=6, OrderDate = DateTime.Parse("2015-04-01"), Region = "East", Rep = "Jones", Item = "Binder", Units = 60, UnitCost = 4.99, Total = 299.4},
                new {Id=7, OrderDate = DateTime.Parse("2015-04-18"), Region = "Central", Rep = "Andrews", Item = "Pencil", Units = 75, UnitCost = 1.99, Total = 149.25},
                new {Id=8, OrderDate = DateTime.Parse("2015-05-05"), Region = "Central", Rep = "Jardine", Item = "Pencil", Units = 90, UnitCost = 4.99, Total = 449.1},
                new {Id=9, OrderDate = DateTime.Parse("2015-05-22"), Region = "West", Rep = "Thompson", Item = "Pencil", Units = 32, UnitCost = 1.99, Total = 63.68},
                new {Id=10, OrderDate = DateTime.Parse("2015-06-08"), Region = "East", Rep = "Jones", Item = "Binder", Units = 60, UnitCost = 8.99, Total = 539.4},
                new {Id=11, OrderDate = DateTime.Parse("2015-06-25"), Region = "Central", Rep = "Morgan", Item = "Pencil", Units = 90, UnitCost = 4.99, Total = 449.1},
                new {Id=12, OrderDate = DateTime.Parse("2015-07-12"), Region = "East", Rep = "Howard", Item = "Binder", Units = 29, UnitCost = 1.99, Total = 57.71},
                new {Id=13, OrderDate = DateTime.Parse("2015-07-29"), Region = "East", Rep = "Parent", Item = "Binder", Units = 81, UnitCost = 19.99, Total = 1619.19},
                new {Id=14, OrderDate = DateTime.Parse("2015-08-15"), Region = "East", Rep = "Jones", Item = "Pencil", Units = 35, UnitCost = 4.99, Total = 174.65},
                new {Id=15, OrderDate = DateTime.Parse("2015-09-01"), Region = "Central", Rep = "Smith", Item = "Desk", Units = 2, UnitCost = 125, Total = 250},
                new {Id=16, OrderDate = DateTime.Parse("2015-09-18"), Region = "East", Rep = "Jones", Item = "Pen Set", Units = 16, UnitCost = 15.99, Total = 255.84},
                new {Id=17, OrderDate = DateTime.Parse("2015-10-05"), Region = "Central", Rep = "Morgan", Item = "Binder", Units = 28, UnitCost = 8.99, Total = 251.72},
                new {Id=18, OrderDate = DateTime.Parse("2015-10-22"), Region = "East", Rep = "Jones", Item = "Pen", Units = 64, UnitCost = 8.99, Total = 575.36},
                new {Id=19, OrderDate = DateTime.Parse("2015-11-08"), Region = "East", Rep = "Parent", Item = "Pen", Units = 15, UnitCost = 19.99, Total = 299.85},
                new {Id=20, OrderDate = DateTime.Parse("2015-11-25"), Region = "Central", Rep = "Kivell", Item = "Pen Set", Units = 96, UnitCost = 4.99, Total = 479.04},
                new {Id=21, OrderDate = DateTime.Parse("2015-12-12"), Region = "Central", Rep = "Smith", Item = "Pencil", Units = 67, UnitCost = 1.29, Total = 86.43},
                new {Id=22, OrderDate = DateTime.Parse("2015-12-29"), Region = "East", Rep = "Parent", Item = "Pen Set", Units = 74, UnitCost = 15.99, Total = 1183.26},
                new {Id=23, OrderDate = DateTime.Parse("2016-01-15"), Region = "Central", Rep = "Gill", Item = "Binder", Units = 46, UnitCost = 8.99, Total = 413.54},
                new {Id=24, OrderDate = DateTime.Parse("2016-02-01"), Region = "Central", Rep = "Smith", Item = "Binder", Units = 87, UnitCost = 15, Total = 1305},
                new {Id=25, OrderDate = DateTime.Parse("2016-02-18"), Region = "East", Rep = "Jones", Item = "Binder", Units = 4, UnitCost = 4.99, Total = 19.96},
                new {Id=26, OrderDate = DateTime.Parse("2016-03-07"), Region = "West", Rep = "Sorvino", Item = "Binder", Units = 7, UnitCost = 19.99, Total = 139.93},
                new {Id=27, OrderDate = DateTime.Parse("2016-03-24"), Region = "Central", Rep = "Jardine", Item = "Pen Set", Units = 50, UnitCost = 4.99, Total = 249.5},
                new {Id=28, OrderDate = DateTime.Parse("2016-04-10"), Region = "Central", Rep = "Andrews", Item = "Pencil", Units = 66, UnitCost = 1.99, Total = 131.34},
                new {Id=29, OrderDate = DateTime.Parse("2016-04-27"), Region = "East", Rep = "Howard", Item = "Pen", Units = 96, UnitCost = 4.99, Total = 479.04},
                new {Id=30, OrderDate = DateTime.Parse("2016-05-14"), Region = "Central", Rep = "Gill", Item = "Pencil", Units = 53, UnitCost = 1.29, Total = 68.37},
                new {Id=31, OrderDate = DateTime.Parse("2016-05-31"), Region = "Central", Rep = "Gill", Item = "Binder", Units = 80, UnitCost = 8.99, Total = 719.2},
                new {Id=32, OrderDate = DateTime.Parse("2016-06-17"), Region = "Central", Rep = "Kivell", Item = "Desk", Units = 5, UnitCost = 125, Total = 625},
                new {Id=33, OrderDate = DateTime.Parse("2016-07-04"), Region = "East", Rep = "Jones", Item = "Pen Set", Units = 62, UnitCost = 4.99, Total = 309.38},
                new {Id=34, OrderDate = DateTime.Parse("2016-07-21"), Region = "Central", Rep = "Morgan", Item = "Pen Set", Units = 55, UnitCost = 12.49, Total = 686.95},
                new {Id=35, OrderDate = DateTime.Parse("2016-08-07"), Region = "Central", Rep = "Kivell", Item = "Pen Set", Units = 42, UnitCost = 23.95, Total = 1005.9},
                new {Id=36, OrderDate = DateTime.Parse("2016-08-24"), Region = "West", Rep = "Sorvino", Item = "Desk", Units = 3, UnitCost = 275, Total = 825},
                new {Id=37, OrderDate = DateTime.Parse("2016-09-10"), Region = "Central", Rep = "Gill", Item = "Pencil", Units = 7, UnitCost = 1.29, Total = 9.03},
                new {Id=38, OrderDate = DateTime.Parse("2016-09-27"), Region = "West", Rep = "Sorvino", Item = "Pen", Units = 76, UnitCost = 1.99, Total = 151.24},
                new {Id=39, OrderDate = DateTime.Parse("2016-10-14"), Region = "West", Rep = "Thompson", Item = "Binder", Units = 57, UnitCost = 19.99, Total = 1139.43},
                new {Id=40, OrderDate = DateTime.Parse("2016-10-31"), Region = "Central", Rep = "Andrews", Item = "Pencil", Units = 14, UnitCost = 1.29, Total = 18.06},
                new {Id=41, OrderDate = DateTime.Parse("2016-11-17"), Region = "Central", Rep = "Jardine", Item = "Binder", Units = 11, UnitCost = 4.99, Total = 54.89},
                new {Id=42, OrderDate = DateTime.Parse("2016-12-04"), Region = "Central", Rep = "Jardine", Item = "Binder", Units = 94, UnitCost = 19.99, Total = 1879.06},
                new {Id=43, OrderDate = DateTime.Parse("2016-12-21"), Region = "Central", Rep = "Andrews", Item = "Binder", Units = 28, UnitCost = 4.99, Total = 139.72}
            };

            ws1.FirstCell()
                .CellBelow()
                .CellRight()
                .InsertTable(data, "Table1");

            var ws2 = wb.AddWorksheet("MixedData");
            ws2.FirstCell().InsertData(new object[] { 1, 2.0, "3", 3, new DateTime(2020, 1, 1), true, new TimeSpan(10, 5, 30, 10) });

            return wb;
        }
    }
}
