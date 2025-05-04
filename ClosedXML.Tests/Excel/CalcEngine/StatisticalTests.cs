// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class StatisticalTests
    {
        private const double tolerance = 1e-6;
        private XLWorkbook workbook;
        
        [TearDown]
        public void Cleanup()
        {
            workbook.Dispose();
        }
        

        [Test]
        public void Average()
        {
            double value;
            value = (double)workbook.Evaluate("AVERAGE(-27.5,93.93,64.51,-70.56)");
            Assert.That(value, Is.EqualTo(15.095).Within(tolerance));

            var ws = workbook.Worksheets.First();
            value = (double)ws.Evaluate("AVERAGE(G3:G45)");
            Assert.Multiple(() =>
            {
                Assert.That(value, Is.EqualTo(49.3255814).Within(tolerance));

                // Column D contains only strings - no average, because non-number types are skipped
                Assert.That(ws.Evaluate("AVERAGE(D3:D45)"), Is.EqualTo(XLError.DivisionByZero));

                // Non-numbers in array are skipped instead of being converted
                Assert.That(ws.Evaluate("AVERAGE({FALSE, TRUE, \"1\", \"0 0/2\", -1})"), Is.EqualTo(-1));
            });

            // Blank value in references are skipped
            ws.Cell("Z1").Value = Blank.Value;
            Assert.That(ws.Evaluate("AVERAGE(Z1,1)"), Is.EqualTo(1));

            AssertScalarToNumberConversion("AVERAGE", 0.5);
            AssertAnyErrorIsPropagated("AVERAGE");
        }

        [Test]
        public void AverageA()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // Examples from specification
            ws.Cell("E1").Value = Blank.Value;
            Assert.That(ws.Evaluate("AVERAGEA(10, E1)"), Is.EqualTo(10));
            ws.Cell("E2").Value = true;
            Assert.That(ws.Evaluate("AVERAGEA(10, E2)"), Is.EqualTo(5.5));
            ws.Cell("E3").Value = false;
            Assert.Multiple(() =>
            {
                Assert.That(ws.Evaluate("AVERAGEA(10, E3)"), Is.EqualTo(5));

                // Make sure multiple values not in an array work as intended
                Assert.That((double)workbook.Evaluate("AVERAGEA(-27.5,93.93,64.51,-70.56)"), Is.EqualTo(15.095).Within(tolerance));

                // Array logical arguments are ignored
                Assert.That(workbook.Evaluate("AVERAGEA({2,TRUE,TRUE,FALSE,FALSE})"), Is.EqualTo(2));

                // Array text arguments are counted as zero (4+2+0+0)/4
                Assert.That(workbook.Evaluate("AVERAGEA({4, 2, \"hello\", \"10\" })"), Is.EqualTo(1.5));
            });

            // Reference argument only counts logical as 0/1, text as 0 and ignores blanks.
            ws.Cell("Z1").Value = Blank.Value; // Not counted
            ws.Cell("Z2").Value = true; // 1
            ws.Cell("Z3").Value = "100"; // 0
            ws.Cell("Z4").Value = "hello"; // 0
            ws.Cell("Z5").Value = 0; // 0
            ws.Cell("Z6").Value = 4; // 4
            Assert.That((double)ws.Evaluate("AVERAGEA(Z1:Z6)"), Is.EqualTo(1));

            AssertScalarToNumberConversion("AVERAGEA", 0.5);
            AssertAnyErrorIsPropagated("AVERAGEA");
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
            Assert.That(result, Is.EqualTo(expected).Within(tolerance));
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
            Assert.That(result, Is.EqualTo(expected).Within(tolerance));
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
            Assert.That(result, Is.EqualTo(XLError.NumberInvalid));
        }

        [Test]
        public void Count()
        {
            var ws = workbook.Worksheets.First();
            XLCellValue value;
            value = ws.Evaluate("COUNT(D3:D45)");
            Assert.That(value, Is.EqualTo(0));

            value = ws.Evaluate("COUNT(G3:G45)");
            Assert.That(value, Is.EqualTo(43));

            value = ws.Evaluate("COUNT(G:G)");
            Assert.That(value, Is.EqualTo(43));

            value = workbook.Evaluate("COUNT(Data!G:G)");
            Assert.Multiple(() =>
            {
                Assert.That(value, Is.EqualTo(43));

                // Scalar blank, logical and text is counted as numbers
                Assert.That(ws.Evaluate("COUNT(IF(TRUE,,),TRUE, FALSE, \"1\")"), Is.EqualTo(4));

                // Non-number values in arrays are not counted as numbers.
                Assert.That(ws.Evaluate("COUNT({TRUE,FALSE,\"1\"})"), Is.EqualTo(0));

                // Text is not counted as number.
                Assert.That(ws.Evaluate("COUNT(\"Hello\")"), Is.EqualTo(0));
            });

            // Blank cells are not counted as numbers
            ws.Cell("Z1").Value = Blank.Value;
            Assert.Multiple(() =>
            {
                Assert.That(ws.Evaluate("COUNT(Z1)"), Is.EqualTo(0));

                // Scalar errors are not propagated
                Assert.That(ws.Evaluate("COUNT(1, #NULL!)"), Is.EqualTo(1));

                // Array errors are not propagated
                Assert.That(ws.Evaluate("COUNT({1, #NULL!})"), Is.EqualTo(1));
            });

            // Reference errors are not propagated
            ws.Cell("Z1").Value = XLError.NullValue;
            Assert.That(ws.Evaluate("COUNT(Z1)"), Is.EqualTo(0));
        }

        [Test]
        public void CountA()
        {
            var ws = workbook.Worksheets.First();
            var value = ws.Evaluate("COUNTA(D3:D45)");
            Assert.That(value, Is.EqualTo(43));

            value = ws.Evaluate("COUNTA(G3:G45)");
            Assert.That(value, Is.EqualTo(43));

            value = ws.Evaluate("COUNTA(G:G)");
            Assert.That(value, Is.EqualTo(44));

            value = workbook.Evaluate("COUNTA(Data!G:G)");
            Assert.That(value, Is.EqualTo(44));
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
            Assert.That(ws.Cell("A9").Value, Is.EqualTo(7));
        }

        [Test]
        public void CountA_on_examples_from_spec()
        {
            Assert.That(XLWorkbook.EvaluateExpr("COUNTA(1,2,3,4,5)"), Is.EqualTo(5));
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr("COUNTA(1,2,3,4,5)"), Is.EqualTo(5));
                Assert.That(XLWorkbook.EvaluateExpr("COUNTA({1,2,3,4,5},6,\"7\")"), Is.EqualTo(7));
            });

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("E2").Value = true;
            Assert.Multiple(() =>
            {
                Assert.That(ws.Evaluate("COUNTA(10, E1)"), Is.EqualTo(1));
                Assert.That(ws.Evaluate("COUNTA(10, E2)"), Is.EqualTo(2));
            });
        }

        [Test]
        public void CountA_accepts_union_references()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A2").Value = 7;
            ws.Cell("B5").Value = false;
            Assert.That(ws.Evaluate("COUNTA((A1:A4,B4:B7))"), Is.EqualTo(2));
        }

        [Test]
        public void CountA_doesnt_count_single_blank_cell_reference()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.That(ws.Evaluate("COUNTA(A1)"), Is.EqualTo(0));
        }

        [Test]
        public void CountA_counts_blank_argument()
        {
            Assert.That(XLWorkbook.EvaluateExpr("COUNTA(IF(TRUE,,))"), Is.EqualTo(1));
        }

        [Test]
        public void CountA_counts_error_arguments()
        {
            Assert.That(XLWorkbook.EvaluateExpr("COUNTA(#NULL!, #DIV/0!, #VALUE!, #REF!, #NAME?, #NUM!, #N/A)"), Is.EqualTo(7));
        }

        [Test]
        public void CountA_counts_empty_string()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = string.Empty;
            Assert.That(ws.Evaluate("COUNTA(A1, \"\")"), Is.EqualTo(2));
        }

        [Test]
        public void CountBlank()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = Blank.Value;
            ws.Cell("A2").Value = 0;
            ws.Cell("A3").Value = 1;
            ws.Cell("A4").Value = false;
            ws.Cell("A5").Value = true;
            ws.Cell("A6").Value = "";
            ws.Cell("A7").Value = "Text";
            ws.Cell("A8").Value = XLError.DivisionByZero;

            Assert.Multiple(() =>
            {
                // Blank and empty text value is counted as blank
                Assert.That(ws.Evaluate("COUNTBLANK(A1)"), Is.EqualTo(1));
                Assert.That(ws.Cell("A6").Value, Is.EqualTo(""));
                Assert.That(ws.Evaluate("COUNTBLANK(A6)"), Is.EqualTo(1));

                // Anything else isn't counted as blank
                Assert.That(ws.Evaluate("COUNTBLANK(A1:A8)"), Is.EqualTo(2));

                Assert.That(ws.Evaluate("COUNTBLANK(A:XFD)"), Is.EqualTo(17179869178d));

                // Check that all others argument types. The Excel grammar doesn't allow that,
                // so use IF workaround for that.
                Assert.That(ws.Evaluate("COUNTBLANK(IF(TRUE,))"), Is.EqualTo(XLError.IncompatibleValue)); // Blank
                Assert.That(ws.Evaluate("COUNTBLANK(IF(TRUE,FALSE))"), Is.EqualTo(XLError.IncompatibleValue)); // Logical
                Assert.That(ws.Evaluate("COUNTBLANK(IF(TRUE,1))"), Is.EqualTo(XLError.IncompatibleValue)); // Number
                Assert.That(ws.Evaluate("COUNTBLANK(IF(TRUE,\"\"))"), Is.EqualTo(XLError.IncompatibleValue)); // Text
                Assert.That(ws.Evaluate("COUNTBLANK(IF(TRUE,#DIV/0!))"), Is.EqualTo(XLError.DivisionByZero)); // Error
                Assert.That(ws.Evaluate("COUNTBLANK(IF(TRUE,{1}))"), Is.EqualTo(XLError.IncompatibleValue)); // Array
            });
        }

        [Test]
        public void CountIf()
        {
            var ws = workbook.Worksheets.First();
            XLCellValue value;
            value = ws.Evaluate(@"=COUNTIF(D3:D45,""Central"")");
            Assert.That(value, Is.EqualTo(24));

            value = ws.Evaluate(@"=COUNTIF(D:D,""Central"")");
            Assert.That(value, Is.EqualTo(24));

            value = workbook.Evaluate(@"=COUNTIF(Data!D:D,""Central"")");
            Assert.That(value, Is.EqualTo(24));
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
            Assert.That(value, Is.EqualTo(expectedResult));
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
            Assert.That(ws.Evaluate(formula), Is.EqualTo(expected));
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
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            ws.Cell(1, 1).Value = cellContent;

            Assert.That((double)ws.Evaluate(formula), Is.EqualTo(expectedResult));
        }

        [TestCase("=COUNTIFS(B1:D1, \"=Yes\")", 1)]
        [TestCase("=COUNTIFS(B1:B4, \"=Yes\", C1:C4, \"=Yes\")", 2)]
        [TestCase("=COUNTIFS(B4:D4, \"=Yes\", B2:D2, \"=Yes\")", 1)]
        public void CountIfs_ReferenceExample1FromExcelDocumentations(
            string formula,
            int expectedOutcome)
        {
            using var wb = new XLWorkbook();
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

            Assert.That(ws.Evaluate(formula), Is.EqualTo(expectedOutcome));
        }

        [Test]
        public void CountIfs_SingleCondition()
        {
            var ws = workbook.Worksheets.First();
            XLCellValue value;
            value = ws.Evaluate(@"=COUNTIFS(D3:D45,""Central"")");
            Assert.That(value, Is.EqualTo(24));

            value = ws.Evaluate(@"=COUNTIFS(D:D,""Central"")");
            Assert.That(value, Is.EqualTo(24));

            value = workbook.Evaluate(@"=COUNTIFS(Data!D:D,""Central"")");
            Assert.That(value, Is.EqualTo(24));
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
            Assert.That(value, Is.EqualTo(expectedResult));
        }

        [TestCase("COUNTIFS(H1:I3, 1, D1:F2, 2)")]
        [TestCase("COUNTIFS(A:B, \"A*\", C:C, \">2\")")]
        public void CountIfs_returns_error_when_areas_dimensions_are_different(string formula)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.That(ws.Evaluate(formula), Is.EqualTo(XLError.IncompatibleValue));
        }

        [OneTimeTearDown]
        public void Dispose()
        {
            workbook.Dispose();
        }

        [TestCase("H3:H45", ExpectedResult = 7.51126069234216)]
        [TestCase("H:H", ExpectedResult = 7.51126069234216)]
        [TestCase("Data!H:H", ExpectedResult = 7.51126069234216)]
        [TestCase("H3:H10", ExpectedResult = 5.26214814727941)]
        [TestCase("H3:H20", ExpectedResult = 7.01281435054797)]
        [TestCase("H3:H30", ExpectedResult = 7.00137389296182)]
        [TestCase("H3:H3", ExpectedResult = 1.99)]
        [TestCase("H10:H20", ExpectedResult = 8.37855107505682)]
        [TestCase("H15:H20", ExpectedResult = 15.8927310267677)]
        [TestCase("H20:H30", ExpectedResult = 7.14321227391814)]
        [DefaultFloatingPointTolerance(1e-12)]
        public double Geomean_calculation(string sourceValue)
        {
            return (double)workbook.Worksheets.First().Evaluate($"GEOMEAN({sourceValue})");
        }

        [TestCase("D3:D45", ExpectedResult = XLError.NumberInvalid)]
        [TestCase("-1, 0, 3", ExpectedResult = XLError.NumberInvalid)]
        [TestCase("0", ExpectedResult = XLError.NumberInvalid)]
        public XLError Geomean_IncorrectCases(string sourceValue)
        {
            var ws = workbook.Worksheets.First();

            return (XLError)ws.Evaluate($"GEOMEAN({sourceValue})");
        }

        [Test]
        [DefaultFloatingPointTolerance(1e-8)]
        public void Geomean()
        {
            Assert.Multiple(() =>
            {
                // Example from the specification
                Assert.That((double)XLWorkbook.EvaluateExpr("GEOMEAN(10.5,5.3,2.9)"), Is.EqualTo(5.4444547024966));
                Assert.That((double)XLWorkbook.EvaluateExpr("GEOMEAN(10.5,{5.3,2.9},\"12\")"), Is.EqualTo(6.6337805880630));

                // GEOMEAN isn't limited by double scale, i.e. it doesn't use naive algorithm for large number.
                Assert.That((double)XLWorkbook.EvaluateExpr("GEOMEAN(1E+307, 1E+307)"), Is.EqualTo(1.0000000000000231E+307d));

                // Scalar blank is counted as a 0
                Assert.That(XLWorkbook.EvaluateExpr("GEOMEAN(IF(TRUE,), 1)"), Is.EqualTo(XLError.NumberInvalid));

                // Scalar logical and text is converted to numbers
                Assert.That((double)XLWorkbook.EvaluateExpr("GEOMEAN(TRUE, \"5\")"), Is.EqualTo(2.236067977));

                // Non-number values in arrays are ignored.
                Assert.That((double)XLWorkbook.EvaluateExpr("GEOMEAN({TRUE, FALSE, \"1\", 7}, 5)"), Is.EqualTo(5.916079783));

                // Scalar non-number text causes an error due to conversion.
                Assert.That(XLWorkbook.EvaluateExpr("GEOMEAN(\"Hello\", 5)"), Is.EqualTo(XLError.IncompatibleValue));
            });

            // Reference non-number arguments are ignored
            var ws = workbook.Worksheets.First();
            ws.Cell("Z1").Value = Blank.Value;
            ws.Cell("Z2").Value = "1";
            ws.Cell("Z3").Value = "hello";
            ws.Cell("Z4").Value = false;
            ws.Cell("Z5").Value = true;
            ws.Cell("Z6").Value = 5;
            Assert.That((double)ws.Evaluate("GEOMEAN(Z1:Z6)"), Is.EqualTo(5));

            AssertAnyErrorIsPropagated("GEOMEAN");
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
            return (double)workbook.Worksheets.First().Evaluate($"DEVSQ({sourceValue})");
        }

        [TestCase("D3:D45", ExpectedResult = XLError.NumberInvalid)]
        public XLError Devsq_IncorrectCases(string sourceValue)
        {
            var ws = workbook.Worksheets.First();

            return (XLError)ws.Evaluate($"DEVSQ({sourceValue})");
        }

        [Test]
        [DefaultFloatingPointTolerance(1e-10)]
        public void Devsq_is_calculated_from_numbers()
        {
            Assert.Multiple(() =>
            {
                Assert.That((double)XLWorkbook.EvaluateExpr("DEVSQ(5.6, 8.2, 9.2)"), Is.EqualTo(6.90666666666666));
                Assert.That((double)XLWorkbook.EvaluateExpr("DEVSQ({ 5.6, 8.2, 9.2})"), Is.EqualTo(6.90666666666666));

                // Array logical arguments are ignored
                Assert.That(workbook.Evaluate("DEVSQ({2,TRUE,TRUE,FALSE,FALSE})"), Is.EqualTo(0));
                Assert.That((double)workbook.Evaluate("DEVSQ({2, 1, 1, 0, 0})"), Is.EqualTo(2.8));

                // Array text arguments are ignored
                Assert.That(workbook.Evaluate("DEVSQ({4, 2, \"hello\", \"10\" })"), Is.EqualTo(2));
            });

            // Non-numerical reference values are ignored.
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = Blank.Value; // Ignored
            ws.Cell("A2").Value = true; // Ignored
            ws.Cell("A3").Value = "100"; // Ignored
            ws.Cell("A4").Value = "hello"; // Ignored
            ws.Cell("A5").Value = 2; // Included
            ws.Cell("A6").Value = 4; // Included
            Assert.That(ws.Evaluate("DEVSQ(A1:A6)"), Is.EqualTo(2));

            AssertScalarToNumberConversion("DEVSQ", 0.5);
            AssertAnyErrorIsPropagated("DEVSQ");
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

        [TestCase("\"asdf\"", ExpectedResult = XLError.IncompatibleValue)]
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
            Assert.That(value, Is.EqualTo(0));

            value = ws.Evaluate(@"=MAX(G3:G45)");
            Assert.That(value, Is.EqualTo(96));

            value = ws.Evaluate(@"=MAX(G:G)");
            Assert.That(value, Is.EqualTo(96));

            value = workbook.Evaluate(@"=MAX(Data!G:G)");
            Assert.That(value, Is.EqualTo(96));

            // Although in most cases blank cells are considered 0, MAX just ignores them.
            value = workbook.Evaluate(@"MAX(-10, Data!X:Z)");
            Assert.That(value, Is.EqualTo(-10));

            // Arrays - numbers are used
            value = workbook.Evaluate(@"MAX(-10, { -6, -5, 7 })");
            Assert.That(value, Is.EqualTo(7));

            // Arrays - non-number and non-error values are skipped.
            value = workbook.Evaluate(@"MAX(-10, { TRUE, FALSE, ""100"" })");
            Assert.That(value, Is.EqualTo(-10));

            // Reference argument ignores everything but number.
            ws.Cell("Z1").Value = Blank.Value;
            ws.Cell("Z2").Value = true;
            ws.Cell("Z3").Value = "100";
            ws.Cell("Z4").Value = "hello";
            ws.Cell("Z5").Value = -4;
            Assert.That(ws.Evaluate("MAX(Z1:Z5)"), Is.EqualTo(-4));

            AssertScalarToNumberConversion("MAX", 1);
            AssertAnyErrorIsPropagated("MAX");
        }

        [Test]
        public void MaxA()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            Assert.Multiple(() =>
            {
                // Examples from specification
                Assert.That(ws.Evaluate("MAXA(10.4,-3.5,12.6)"), Is.EqualTo(12.6));
                Assert.That(ws.Evaluate("MAXA(10.4,{-3.5,12.6})"), Is.EqualTo(12.6));
                Assert.That(ws.Evaluate("MAXA({\"ABC\",TRUE})"), Is.EqualTo(0));
            });
            ws.Cell("B3").Value = Blank.Value;
            Assert.That(ws.Evaluate("MAX(-10,-12,-15,B3)"), Is.EqualTo(-10));
            ws.Cell("B3").Value = 0;
            Assert.Multiple(() =>
            {
                Assert.That(ws.Evaluate("MAXA(-10,-12,-15,B3)"), Is.EqualTo(0));

                // Array logical arguments are ignored
                Assert.That(workbook.Evaluate("MAXA({-2, TRUE, TRUE, FALSE, FALSE})"), Is.EqualTo(-2));

                // Array text arguments are ignored
                Assert.That(workbook.Evaluate("MAXA({-4, -2, \"hello\", \"10\" })"), Is.EqualTo(-2));
            });

            // Reference argument only counts logical as 0/1, text as 0 and ignores blanks.
            ws.Cell("A1").Value = Blank.Value;
            ws.Cell("A2").Value = true;
            ws.Cell("A3").Value = "100";
            ws.Cell("A4").Value = "hello";
            ws.Cell("A5").Value = -4;
            Assert.Multiple(() =>
            {
                Assert.That(ws.Evaluate("MAXA(A1:A5)"), Is.EqualTo(1));
                Assert.That(ws.Evaluate("MAXA(A3:A5)"), Is.EqualTo(0));
            });

            AssertScalarToNumberConversion("MAXA", 1);
            AssertAnyErrorIsPropagated("MAXA");
        }

        [Test]
        public void Median_with_area_without_numeric_values_returns_error()
        {
            var ws = workbook.Worksheets.First();

            // Column D contains names of regions
            Assert.That(ws.Evaluate("MEDIAN(D3:D45)"), Is.EqualTo(XLError.NumberInvalid));
        }

        [Test]
        public void Median_EvenCountOfCellRange_ReturnsAverageOfTwoElementsInMiddleOfSortedList()
        {
            //Arrange
            var ws = workbook.Worksheets.First();

            //Act
            var value = (double)ws.Evaluate("MEDIAN(I3:I10)");

            //Assert
            Assert.That(value, Is.EqualTo(244.225).Within(tolerance));
        }

        [Test]
        public void Median_EvenCountOfManualNumbers_ReturnsAverageOfTwoElementsInMiddleOfSortedList()
        {
            //Act
            var value = (double)workbook.Evaluate("MEDIAN(-27.5,93.93,64.51,-70.56)");

            //Assert
            Assert.That(value, Is.EqualTo(18.505).Within(tolerance));
        }

        [Test]
        public void Median_OddCountOfCellRange_ReturnsElementInMiddleOfSortedList()
        {
            //Arrange
            var ws = workbook.Worksheets.First();

            //Act
            var value = (double)ws.Evaluate("MEDIAN(I3:I11)");

            //Assert
            Assert.That(value, Is.EqualTo(189.05).Within(tolerance));
        }

        [Test]
        public void Median_OddCountOfManualNumbers_ReturnsElementInMiddleOfSortedList()
        {
            //Act
            var value = (double)workbook.Evaluate("MEDIAN(-27.5,93.93,64.51,-70.56,101.65)");

            //Assert
            Assert.That(value, Is.EqualTo(64.51).Within(tolerance));
        }

        [Test]
        public void Median_uses_only_numbers()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            Assert.Multiple(() =>
            {
                // Examples from specification
                Assert.That(ws.Evaluate("MEDIAN(10, 20)"), Is.EqualTo(15));
                Assert.That(ws.Evaluate("MEDIAN(-3.5, 1.4, 6.9, -4.5)"), Is.EqualTo(-1.05));
                Assert.That(ws.Evaluate("MEDIAN({ -3.5,1.4,6.9},-4.5)"), Is.EqualTo(-1.05));
            });

            // Reference with no value will return error
            ws.Cell("A1").Value = Blank.Value;
            Assert.Multiple(() =>
            {
                Assert.That(ws.Evaluate("MEDIAN(A1)"), Is.EqualTo(XLError.NumberInvalid));

                // Array non-number values are ignored
                Assert.That(ws.Evaluate("MEDIAN({7, TRUE,FALSE,\"1\"})"), Is.EqualTo(7));
            });

            // Only numbers are used from reference, rest is ignored
            ws.Cell("A1").Value = Blank.Value;
            ws.Cell("A2").Value = true;
            ws.Cell("A3").Value = "100";
            ws.Cell("A4").Value = "hello";
            ws.Cell("A5").Value = 0;
            ws.Cell("A6").Value = 4;
            ws.Cell("A7").Value = 5;
            Assert.That(ws.Evaluate("MEDIAN(A1:A7)"), Is.EqualTo(4));

            AssertScalarToNumberConversion("MEDIAN", 0.5);
            AssertAnyErrorIsPropagated("MEDIAN");
        }

        [Test]
        public void Min()
        {
            var ws = workbook.Worksheets.First();
            Assert.Multiple(() =>
            {
                Assert.That(ws.Evaluate("MIN(D3:D45)"), Is.EqualTo(0));
                Assert.That(ws.Evaluate("MIN(G3:G45)"), Is.EqualTo(2));
                Assert.That(ws.Evaluate("MIN(G:G)"), Is.EqualTo(2));
                Assert.That(workbook.Evaluate("MIN(Data!G:G)"), Is.EqualTo(2));

                // Array non-number arguments are ignored
                Assert.That(workbook.Evaluate("MIN({5, TRUE, FALSE, \"1\", \"hello\"})"), Is.EqualTo(5));
            });

            // Reference non-number arguments are ignored
            ws.Cell("Z1").Value = Blank.Value;
            ws.Cell("Z2").Value = "1";
            ws.Cell("Z3").Value = "hello";
            ws.Cell("Z4").Value = false;
            ws.Cell("Z5").Value = true;
            ws.Cell("Z6").Value = 5;
            Assert.Multiple(() =>
            {
                Assert.That(ws.Evaluate("MIN(Z1:Z6)"), Is.EqualTo(5));

                // If there is no value, return 0
                Assert.That(ws.Evaluate("MIN({\"hello\"})"), Is.EqualTo(0));
            });

            AssertScalarToNumberConversion("MIN", 0);
            AssertAnyErrorIsPropagated("MIN");
        }

        [Test]
        public void MinA()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            Assert.Multiple(() =>
            {
                // Examples from specification
                Assert.That(ws.Evaluate("MINA(10.4, -3.5, 12.6)"), Is.EqualTo(-3.5));
                Assert.That(ws.Evaluate("MINA(10.4, {-3.5, 12.6})"), Is.EqualTo(-3.5));
                Assert.That(ws.Evaluate("MINA({\"ABC\", TRUE})"), Is.EqualTo(0));
            });
            ws.Cell("B3").Value = Blank.Value;
            Assert.That(ws.Evaluate("MINA(10, 12, 15, B3)"), Is.EqualTo(10));
            ws.Cell("B3").Value = "Text";
            Assert.That(ws.Evaluate("MINA(10, 12, 15, B3)"), Is.EqualTo(0));

            // Blanks in references are ignored and when MINA doesn't have any values, it returns 0
            ws.Cell("A1").Value = Blank.Value;
            Assert.Multiple(() =>
            {
                Assert.That(ws.Evaluate("MINA(A1)"), Is.EqualTo(0));

                // Array logical arguments are ignored
                Assert.That(wb.Evaluate("MINA({2, TRUE, TRUE, FALSE, FALSE})"), Is.EqualTo(2));

                // Array text arguments are ignored
                Assert.That(wb.Evaluate("MINA({4, 2, \"hello\", \"1\"})"), Is.EqualTo(2));
            });

            // Reference argument only counts logical as 0/1, text as 0 and ignores blanks.
            ws.Cell("A1").Value = Blank.Value; // Ignores
            ws.Cell("A2").Value = true; // Includes
            ws.Cell("A3").Value = "100"; // Considers 0
            ws.Cell("A4").Value = "hello"; // Considers 0
            ws.Cell("A5").Value = -4; // Included
            Assert.Multiple(() =>
            {
                Assert.That(ws.Evaluate("MINA(A1:A2)"), Is.EqualTo(1));
                Assert.That(ws.Evaluate("MINA(A1:A3)"), Is.EqualTo(0));
                Assert.That(ws.Evaluate("MINA(A1:A5)"), Is.EqualTo(-4));
            });

            AssertScalarToNumberConversion("MINA", 0);
            AssertAnyErrorIsPropagated("MINA");
        }

        [Test]
        [DefaultFloatingPointTolerance(tolerance)]
        public void StDev()
        {
            var ws = workbook.Worksheets.First();

            // Only non-convertible text in D column, thus less than 2 samples will return error
            Assert.That(ws.Evaluate("STDEV(D3:D45)"), Is.EqualTo(XLError.DivisionByZero));

            // Calculate StDev from numeric values (reference contains only numbers)
            var value = (double)ws.Evaluate("STDEV(H3:H45)");
            Assert.That(value, Is.EqualTo(47.34511769).Within(tolerance));

            // Ignores text values in the H column and only uses numeric ones, same as reference with only number
            value = (double)ws.Evaluate("STDEV(H:H)");
            Assert.That(value, Is.EqualTo(47.34511769).Within(tolerance));

            value = (double)workbook.Evaluate("STDEV(Data!H:H)");
            Assert.Multiple(() =>
            {
                Assert.That(value, Is.EqualTo(47.34511769).Within(tolerance));

                // Need at least two values, otherwise returns error
                Assert.That(workbook.Evaluate("STDEV(1)"), Is.EqualTo(XLError.DivisionByZero));
                Assert.That(workbook.Evaluate("STDEV(0, 0)"), Is.EqualTo(0));

                // Array non-number arguments are ignored
                Assert.That((double)workbook.Evaluate("STDEV({0, 1, \"Hello\", FALSE, TRUE})"), Is.EqualTo(0.707106781).Within(tolerance));
            });

            // Reference argument only uses number, ignores blanks, logical and text
            ws.Cell("Z1").Value = Blank.Value;
            ws.Cell("Z2").Value = true;
            ws.Cell("Z3").Value = "100";
            ws.Cell("Z4").Value = "hello";
            ws.Cell("Z5").Value = 0;
            ws.Cell("Z6").Value = 1;
            Assert.That((double)ws.Evaluate("STDEV(Z1:Z6)"), Is.EqualTo(0.707106781).Within(tolerance));

            AssertScalarToNumberConversion("STDEV", 0.707106781);
            AssertAnyErrorIsPropagated("STDEV");
        }

        [Test]
        [DefaultFloatingPointTolerance(tolerance)]
        public void StDevA()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            Assert.Multiple(() =>
            {
                // Example from specification
                Assert.That((double)ws.Evaluate("STDEVA(123, 134, 143, 173, 112, 109)"), Is.EqualTo(23.72902583));

                // Array non-number arguments are ignored
                Assert.That((double)ws.Evaluate("STDEVA({0, 1, \"9\", \"Hello\", FALSE, TRUE})"), Is.EqualTo(0.707106781));
            });

            // Reference argument ignores blanks, uses numbers, logical and text as zero
            ws.Cell("A1").Value = Blank.Value; // Ignore
            ws.Cell("A2").Value = true; // Include
            ws.Cell("A3").Value = ""; // Consider 0
            ws.Cell("A4").Value = "100"; // Consider 0
            ws.Cell("A5").Value = "hello"; // Consider 0
            ws.Cell("A6").Value = 5;
            ws.Cell("A7").Value = 7;
            Assert.Multiple(() =>
            {
                Assert.That((double)ws.Evaluate("STDEVA(A1:A7)"), Is.EqualTo(3.060501048));

                // Need at least one sample, otherwise returns error (text in array is ignored)
                Assert.That(ws.Evaluate("STDEVA({\"hello\"})"), Is.EqualTo(XLError.DivisionByZero));
            });

            AssertScalarToNumberConversion("STDEVA", 0.707106781);
            AssertAnyErrorIsPropagated("STDEVA");
        }

        [Test]
        public void StDevP()
        {
            var ws = workbook.Worksheets.First();

            Assert.Multiple(() =>
            {
                // Example from specification
                Assert.That((double)ws.Evaluate("STDEVP(123, 134, 143, 173, 112, 109)"), Is.EqualTo(21.66153785).Within(tolerance));

                // Column D contains only region names (non-convertible text), thus reference contains less than 1 sample that is required
                Assert.That(ws.Evaluate("STDEVP(D3:D45)"), Is.EqualTo(XLError.DivisionByZero));

                // Calculate StDevP from numeric values (reference contains only numbers)
                Assert.That((double)ws.Evaluate("STDEVP(H3:H45)"), Is.EqualTo(46.79135458).Within(tolerance));

                // StDevP ignores text values/blanks in the H column and only uses numeric ones, the result is same as the reference above that contains only numbers
                Assert.That((double)ws.Evaluate("STDEVP(H:H)"), Is.EqualTo(46.79135458).Within(tolerance));

                Assert.That((double)workbook.Evaluate("STDEVP(Data!H:H)"), Is.EqualTo(46.79135458).Within(tolerance));

                // If sample size is 0, return error
                Assert.That(workbook.Evaluate("STDEVP({TRUE})"), Is.EqualTo(XLError.DivisionByZero));
                Assert.That(workbook.Evaluate("STDEVP(100)"), Is.EqualTo(0));

                // Array non-number arguments are ignored
                Assert.That(workbook.Evaluate("STDEVP({0, 1, \"Hello\", FALSE, TRUE})"), Is.EqualTo(0.5));
            });

            // Reference argument only uses numbers, ignores blanks, logical and text
            ws.Cell("Z1").Value = Blank.Value;
            ws.Cell("Z2").Value = true;
            ws.Cell("Z3").Value = "100";
            ws.Cell("Z4").Value = "hello";
            ws.Cell("Z5").Value = 0;
            ws.Cell("Z6").Value = 1;
            Assert.That(ws.Evaluate("STDEVP(Z1:Z6)"), Is.EqualTo(0.5));

            AssertScalarToNumberConversion("STDEVP", 0.5);
            AssertAnyErrorIsPropagated("STDEVP");
        }

        [Test]
        [DefaultFloatingPointTolerance(tolerance)]
        public void StDevPA()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            Assert.Multiple(() =>
            {
                // Example from specification
                Assert.That((double)ws.Evaluate("STDEVPA(123, 134, 143, 173, 112, 109)"), Is.EqualTo(21.66153785));

                // Array non-number arguments are ignored
                Assert.That((double)ws.Evaluate("STDEVPA({0, 1, \"9\", \"Hello\", FALSE, TRUE})"), Is.EqualTo(0.5));
            });

            // Reference argument ignores blanks, uses numbers, logical and text as zero
            ws.Cell("A1").Value = Blank.Value; // Ignore
            ws.Cell("A2").Value = true; // Include
            ws.Cell("A3").Value = ""; // Consider 0
            ws.Cell("A4").Value = "100"; // Consider 0
            ws.Cell("A5").Value = "hello"; // Consider 0
            ws.Cell("A6").Value = 5;
            ws.Cell("A7").Value = 7;
            Assert.Multiple(() =>
            {
                Assert.That((double)ws.Evaluate("STDEVPA(A1:A7)"), Is.EqualTo(2.793842436));

                // Need at least one sample, otherwise returns error (text in array is ignored)
                Assert.That(ws.Evaluate("STDEVPA({\"hello\"})"), Is.EqualTo(XLError.DivisionByZero));
            });

            AssertScalarToNumberConversion("STDEVPA", 0.5);
            AssertAnyErrorIsPropagated("STDEVPA");
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
            Assert.That(ws.Evaluate(formula), Is.EqualTo(expected));
        }

        [Test]
        public void SumIf_specification_examples()
        {
            // Test examples from specification.
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 3;
            ws.Cell("B1").Value = 10;
            ws.Cell("C1").Value = 7;
            ws.Cell("D1").Value = 10;

            Assert.Multiple(() =>
            {
                Assert.That(ws.Evaluate("SUMIF(A1:D1,\"=10\")"), Is.EqualTo(20));
                Assert.That(ws.Evaluate("SUMIF(A1:D1,\">5\")"), Is.EqualTo(27));
                Assert.That(ws.Evaluate("SUMIF(A1:D1,\"<>10\")"), Is.EqualTo(10));
            });

            ws.Cell("A2").Value = "apples";
            ws.Cell("B2").Value = "melons";
            ws.Cell("C2").Value = 10;
            ws.Cell("D2").Value = 15;
            Assert.That(ws.Evaluate("SUMIF(A2:B2,\"*es\",C2:D2)"), Is.EqualTo(10));
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
            using var wb = SetupWorkbook();
            var ws = wb.Worksheets.First();
            //Let's pre-initialize cells we need so they didn't affect the result
            ws.Range("A1:J45").Style.Fill.BackgroundColor = XLColor.Amber;
            ws.Cell("ZZ1000").Value = 1;

            var actualResult = (double)ws.Evaluate(formulaA1);

            Assert.That(actualResult, Is.EqualTo(expectedResult).Within(tolerance));
        }

        [Test]
        public void Var()
        {
            var ws = workbook.Worksheets.First();

            Assert.Multiple(() =>
            {
                // Example from specification
                Assert.That(ws.Evaluate("VAR(1202,1220,1323,1254,1302)"), Is.EqualTo(2683.2));

                // Only non-convertible text in D column, thus less than 2 samples.
                Assert.That(ws.Evaluate("VAR(D3:D45)"), Is.EqualTo(XLError.DivisionByZero));

                // Calculate VAR from numeric values (reference contains only numbers)
                Assert.That((double)ws.Evaluate("VAR(H3:H45)"), Is.EqualTo(2241.560169).Within(tolerance));

                // Ignores text values in the H column and only uses numeric ones, same as reference with only number
                Assert.That((double)ws.Evaluate("VAR(H:H)"), Is.EqualTo(2241.560169).Within(tolerance));
                Assert.That((double)workbook.Evaluate("VAR(Data!H:H)"), Is.EqualTo(2241.560169).Within(tolerance));

                // Need at least two samples, otherwise returns error
                Assert.That(workbook.Evaluate("VAR({\"hello\"})"), Is.EqualTo(XLError.DivisionByZero));
                Assert.That(workbook.Evaluate("VAR(5)"), Is.EqualTo(XLError.DivisionByZero));
                Assert.That(workbook.Evaluate("VAR(5, 6)"), Is.EqualTo(0.5));

                // Array non-number arguments are ignored
                Assert.That(workbook.Evaluate("VAR({0, 1, \"Hello\", FALSE, TRUE})"), Is.EqualTo(0.5));
            });

            // Reference argument only uses number, ignores blanks, logical and text
            ws.Cell("Z1").Value = Blank.Value;
            ws.Cell("Z2").Value = true;
            ws.Cell("Z3").Value = "100";
            ws.Cell("Z4").Value = "hello";
            ws.Cell("Z5").Value = 0;
            ws.Cell("Z6").Value = 1;
            Assert.That(ws.Evaluate("VAR(Z1:Z6)"), Is.EqualTo(0.5));

            AssertScalarToNumberConversion("VAR", 0.5);
            AssertAnyErrorIsPropagated("VAR");
        }

        [Test]
        [DefaultFloatingPointTolerance(tolerance)]
        public void VarA()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            Assert.Multiple(() =>
            {
                // Example from specification
                Assert.That(ws.Evaluate("VARA(1202, 1220, 1323, 1254, 1302)"), Is.EqualTo(2683.2));

                // Array non-number arguments are ignored
                Assert.That(ws.Evaluate("VARA({5, 7, \"9\", \"Hello\", FALSE, TRUE})"), Is.EqualTo(2));
            });

            // Reference argument ignores blanks, uses numbers, logical and text as zero
            ws.Cell("A1").Value = Blank.Value; // Ignore
            ws.Cell("A2").Value = true; // Include
            ws.Cell("A3").Value = ""; // Consider 0
            ws.Cell("A4").Value = "100"; // Consider 0
            ws.Cell("A5").Value = "hello"; // Consider 0
            ws.Cell("A6").Value = 5;
            ws.Cell("A7").Value = 7;
            Assert.Multiple(() =>
            {
                Assert.That((double)ws.Evaluate("VARA(A1:A7)"), Is.EqualTo(9.366666667));

                // Need at least one sample, otherwise returns error (text in array is ignored)
                Assert.That(ws.Evaluate("VARA({\"hello\"})"), Is.EqualTo(XLError.DivisionByZero));
            });

            AssertScalarToNumberConversion("VARA", 0.5);
            AssertAnyErrorIsPropagated("VARA");
        }

        [Test]
        public void VarP()
        {
            var ws = workbook.Worksheets.First();

            Assert.Multiple(() =>
            {
                // Example from specification
                Assert.That((double)ws.Evaluate("VARP(1202,1220,1323,1254,1302)"), Is.EqualTo(2146.56).Within(tolerance));

                // Only non-convertible text in D column, thus less than 1 sample.
                Assert.That(ws.Evaluate("VARP(D3:D45)"), Is.EqualTo(XLError.DivisionByZero));

                // Calculate VARP from numeric values (reference contains only numbers)
                Assert.That((double)ws.Evaluate("VARP(H3:H45)"), Is.EqualTo(2189.430863).Within(tolerance));

                // Ignores text values in the H column and only uses numeric ones, same as reference with only number
                Assert.That((double)ws.Evaluate("VARP(H:H)"), Is.EqualTo(2189.430863).Within(tolerance));
                Assert.That((double)workbook.Evaluate("VARP(Data!H:H)"), Is.EqualTo(2189.430863).Within(tolerance));

                // Need at least one sample, otherwise returns error
                Assert.That(workbook.Evaluate("VARP({\"hello\"})"), Is.EqualTo(XLError.DivisionByZero));
                Assert.That(workbook.Evaluate("VARP(5)"), Is.EqualTo(0));

                // Array non-number arguments are ignored
                Assert.That(workbook.Evaluate("VARP({0, 1, \"Hello\", FALSE, TRUE})"), Is.EqualTo(0.25));
            });

            // Reference argument only uses number, ignores blanks, logical and text
            ws.Cell("Z1").Value = Blank.Value;
            ws.Cell("Z2").Value = true;
            ws.Cell("Z3").Value = "100";
            ws.Cell("Z4").Value = "hello";
            ws.Cell("Z5").Value = 0;
            ws.Cell("Z6").Value = 1;
            Assert.That(ws.Evaluate("VARP(Z1:Z6)"), Is.EqualTo(0.25));

            AssertScalarToNumberConversion("VARP", 0.25);
            AssertAnyErrorIsPropagated("VARP");
        }

        [Test]
        [DefaultFloatingPointTolerance(tolerance)]
        public void VarPA()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            Assert.Multiple(() =>
            {
                // Example from specification
                Assert.That(ws.Evaluate("VARPA(1202, 1220, 1323, 1254, 1302)"), Is.EqualTo(2146.56));

                // Array non-number arguments are ignored
                Assert.That(ws.Evaluate("VARPA({5, 7, \"9\", \"Hello\", FALSE, TRUE})"), Is.EqualTo(1));
            });

            // Reference argument ignores blanks, uses numbers, logical and text as zero
            ws.Cell("A1").Value = Blank.Value; // Ignore
            ws.Cell("A2").Value = true; // Include
            ws.Cell("A3").Value = ""; // Consider 0
            ws.Cell("A4").Value = "100"; // Consider 0
            ws.Cell("A5").Value = "hello"; // Consider 0
            ws.Cell("A6").Value = 5;
            ws.Cell("A7").Value = 7;
            Assert.Multiple(() =>
            {
                Assert.That((double)ws.Evaluate("VARPA(A1:A7)"), Is.EqualTo(7.805555556));

                // Need at least one sample, otherwise returns error (text in array is ignored)
                Assert.That(ws.Evaluate("VARPA({\"hello\"})"), Is.EqualTo(XLError.DivisionByZero));
            });

            AssertScalarToNumberConversion("VARPA", 0.25);
            AssertAnyErrorIsPropagated("VARPA");
        }

        [Test]
        public void Large()
        {
            var ws = workbook.Worksheet("Data");
            var value = ws.Evaluate("LARGE(G1:G45, 1)");
            Assert.That(value, Is.EqualTo(96));

            value = ws.Evaluate("LARGE(G1:G45, 7)");
            Assert.That(value, Is.EqualTo(87));

            value = ws.Evaluate("LARGE(G1:G45, 0)");
            Assert.That(value, Is.EqualTo(XLError.NumberInvalid));

            value = ws.Evaluate("LARGE(G1:G45, -1)");
            Assert.That(value, Is.EqualTo(XLError.NumberInvalid));

            value = ws.Evaluate("LARGE(G1:G45,\"test\")");
            Assert.That(value, Is.EqualTo(XLError.IncompatibleValue));

            value = ws.Evaluate("LARGE(C:C,7)");
            Assert.That(value, Is.EqualTo(42623));

            value = ws.Evaluate("LARGE(D:D,7)");
            Assert.That(value, Is.EqualTo(XLError.NumberInvalid));

            ws = workbook.Worksheet("MixedData");

            value = ws.Evaluate("LARGE(A1:A7,6)");
            Assert.That(value, Is.EqualTo(XLError.NumberInvalid));

            // Ignores non-numbers.
            value = ws.Evaluate("LARGE(A1:A7,5)");
            Assert.That(value, Is.EqualTo(1));

            // Accepts non-area references.
            value = ws.Evaluate("LARGE((A1:A2,A4:A6),2)");
            Assert.That(value, Is.EqualTo(3));

            // Errors are returned.
            value = ws.Evaluate("LARGE({ 1, 2, #N/A }, 1)");
            Assert.That(value, Is.EqualTo(XLError.NoValueAvailable));

            // Uses ceiling logic for number (1.1 -> 2) + can use arrays.
            value = ws.Evaluate("LARGE({ 1, 2 }, 1.1)");
            Assert.That(value, Is.EqualTo(1));

            // If a scalar number-like value supplied, it is converted to number.
            value = ws.Evaluate("LARGE(\"1 1/2\", 1)");
            Assert.That(value, Is.EqualTo(1.5));

            // When the scalar can't be converted, return conversion error.
            value = ws.Evaluate("LARGE(\"test\", 1)");
            Assert.That(value, Is.EqualTo(XLError.IncompatibleValue));
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

        private static void AssertScalarToNumberConversion(string functionName, double result)
        {
            Assert.Multiple(() =>
            {
                // Scalar blank is converted to 0
                Assert.That((double)XLWorkbook.EvaluateExpr($"{functionName}(IF(TRUE,), 1)"), Is.EqualTo(result));

                // Scalar logical is converted to a number
                Assert.That((double)XLWorkbook.EvaluateExpr($"{functionName}(FALSE, TRUE)"), Is.EqualTo(result));
                Assert.That((double)XLWorkbook.EvaluateExpr($"{functionName}(0, TRUE)"), Is.EqualTo(result));
                Assert.That((double)XLWorkbook.EvaluateExpr($"{functionName}(FALSE, 1)"), Is.EqualTo(result));

                // Scalar text is converted to a number
                Assert.That((double)XLWorkbook.EvaluateExpr($"{functionName}(\"0\", \"1\")"), Is.EqualTo(result));
                Assert.That((double)XLWorkbook.EvaluateExpr($"{functionName}(\"1\", \"0 0/2\")"), Is.EqualTo(result));

                // Scalar text that is not convertible returns error
                Assert.That(XLWorkbook.EvaluateExpr($"{functionName}(5, \"Hello\")"), Is.EqualTo(XLError.IncompatibleValue));
            });
        }

        /// <summary>
        /// Assert that a function propagates any error, whether from scalar, array or reference argument.
        /// </summary>
        /// <param name="functionName">Name of a function that accepts any value as argument.</param>
        private static void AssertAnyErrorIsPropagated(string functionName)
        {
            Assert.Multiple(() =>
            {
                // Scalar error is propagated
                Assert.That(XLWorkbook.EvaluateExpr($"{functionName}(1, #NULL!)"), Is.EqualTo(XLError.NullValue));

                // Array error is propagated
                Assert.That(XLWorkbook.EvaluateExpr($"{functionName}({{1, #NULL!}})"), Is.EqualTo(XLError.NullValue));
            });

            // Reference error is propagated
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B1").Value = XLError.NoValueAvailable;
            ws.Cell("B2").Value = 1;
            Assert.Multiple(() =>
            {
                Assert.That(ws.Evaluate($"{functionName}(B1)"), Is.EqualTo(XLError.NoValueAvailable));
                Assert.That(ws.Evaluate($"{functionName}(B1:B2)"), Is.EqualTo(XLError.NoValueAvailable));
            });
        }
    }
}
