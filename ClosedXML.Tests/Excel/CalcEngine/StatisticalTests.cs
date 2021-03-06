// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;
using System;
using System.Linq;

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
            value = workbook.Evaluate("AVERAGE(-27.5,93.93,64.51,-70.56)").CastTo<double>();
            Assert.AreEqual(15.095, value, tolerance);

            var ws = workbook.Worksheets.First();
            value = ws.Evaluate("AVERAGE(G3:G45)").CastTo<double>();
            Assert.AreEqual(49.3255814, value, tolerance);

            Assert.That(() => ws.Evaluate("AVERAGE(D3:D45)"), Throws.TypeOf<ApplicationException>());
        }

        [Test]
        public void Count()
        {
            var ws = workbook.Worksheets.First();
            int value;
            value = ws.Evaluate(@"=COUNT(D3:D45)").CastTo<int>();
            Assert.AreEqual(0, value);

            value = ws.Evaluate(@"=COUNT(G3:G45)").CastTo<int>();
            Assert.AreEqual(43, value);

            value = ws.Evaluate(@"=COUNT(G:G)").CastTo<int>();
            Assert.AreEqual(43, value);

            value = workbook.Evaluate(@"=COUNT(Data!G:G)").CastTo<int>();
            Assert.AreEqual(43, value);
        }

        [Test]
        public void CountA()
        {
            var ws = workbook.Worksheets.First();
            int value;
            value = ws.Evaluate(@"=COUNTA(D3:D45)").CastTo<int>();
            Assert.AreEqual(43, value);

            value = ws.Evaluate(@"=COUNTA(G3:G45)").CastTo<int>();
            Assert.AreEqual(43, value);

            value = ws.Evaluate(@"=COUNTA(G:G)").CastTo<int>();
            Assert.AreEqual(44, value);

            value = workbook.Evaluate(@"=COUNTA(Data!G:G)").CastTo<int>();
            Assert.AreEqual(44, value);
        }

        [Test]
        public void CountBlank()
        {
            var ws = workbook.Worksheets.First();
            int value;
            value = ws.Evaluate(@"=COUNTBLANK(B:B)").CastTo<int>();
            Assert.AreEqual(1048532, value);

            value = ws.Evaluate(@"=COUNTBLANK(D43:D49)").CastTo<int>();
            Assert.AreEqual(4, value);

            value = ws.Evaluate(@"=COUNTBLANK(E3:E45)").CastTo<int>();
            Assert.AreEqual(0, value);

            value = ws.Evaluate(@"=COUNTBLANK(A1)").CastTo<int>();
            Assert.AreEqual(1, value);

            Assert.AreEqual(XLCalculationErrorType.NoValueAvailable, workbook.Evaluate(@"=COUNTBLANK(E3:E45)"));
            Assert.Throws<ExpressionParseException>(() => ws.Evaluate(@"=COUNTBLANK()"));
            Assert.Throws<ExpressionParseException>(() => ws.Evaluate(@"=COUNTBLANK(A3:A45,E3:E45)"));
        }

        [Test]
        public void CountIf()
        {
            var ws = workbook.Worksheets.First();
            int value;
            value = ws.Evaluate(@"=COUNTIF(D3:D45,""Central"")").CastTo<int>();
            Assert.AreEqual(24, value);

            value = ws.Evaluate(@"=COUNTIF(D:D,""Central"")").CastTo<int>();
            Assert.AreEqual(24, value);

            value = workbook.Evaluate(@"=COUNTIF(Data!D:D,""Central"")").CastTo<int>();
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

            int value = ws.Evaluate(formula).CastTo<int>();
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
            int value;
            value = ws.Evaluate(@"=COUNTIFS(D3:D45,""Central"")").CastTo<int>();
            Assert.AreEqual(24, value);

            value = ws.Evaluate(@"=COUNTIFS(D:D,""Central"")").CastTo<int>();
            Assert.AreEqual(24, value);

            value = workbook.Evaluate(@"=COUNTIFS(Data!D:D,""Central"")").CastTo<int>();
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

            int value = ws.Evaluate(formula).CastTo<int>();
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
            return workbook.Worksheets.First().Evaluate($"=GEOMEAN({sourceValue})").CastTo<double>();
        }

        [TestCase("D3:D45", ExpectedResult = XLCalculationErrorType.NumberInvalid)]
        [TestCase("-1, 0, 3", ExpectedResult = XLCalculationErrorType.NumberInvalid)]
        public XLCalculationErrorType Geomean_IncorrectCases(string sourceValue)
        {
            var ws = workbook.Worksheets.First();

            return (XLCalculationErrorType)ws.Evaluate($"=GEOMEAN({sourceValue})");
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
            return workbook.Worksheets.First().Evaluate($"=DEVSQ({sourceValue})").CastTo<double>();
        }

        [TestCase("D3:D45", ExpectedResult = XLCalculationErrorType.CellValue)]
        public XLCalculationErrorType Devsq_IncorrectCases(string sourceValue)
        {
            var ws = workbook.Worksheets.First();

            return (XLCalculationErrorType)ws.Evaluate($"=DEVSQ({sourceValue})");
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
            return XLWorkbook.EvaluateExpr($"=FISHER({sourceValue})").CastTo<double>();
        }

        // TODO : the string case will be treated correctly when Coercion is implemented better
        //[TestCase("asdf", typeof(CellValueException), "Parameter non numeric.")]
        [TestCase("5", ExpectedResult = XLCalculationErrorType.NumberInvalid)]
        [TestCase("-1", ExpectedResult = XLCalculationErrorType.NumberInvalid)]
        [TestCase("1", ExpectedResult = XLCalculationErrorType.NumberInvalid)]
        public XLCalculationErrorType Fisher_IncorrectCases(string sourceValue)
        {
            return (XLCalculationErrorType)XLWorkbook.EvaluateExpr($"=FISHER({sourceValue})");
        }

        [Test]
        public void Max()
        {
            var ws = workbook.Worksheets.First();
            int value;
            value = ws.Evaluate(@"=MAX(D3:D45)").CastTo<int>();
            Assert.AreEqual(0, value);

            value = ws.Evaluate(@"=MAX(G3:G45)").CastTo<int>();
            Assert.AreEqual(96, value);

            value = ws.Evaluate(@"=MAX(G:G)").CastTo<int>();
            Assert.AreEqual(96, value);

            value = workbook.Evaluate(@"=MAX(Data!G:G)").CastTo<int>();
            Assert.AreEqual(96, value);
        }

        [Test]
        public void Min()
        {
            var ws = workbook.Worksheets.First();
            int value;
            value = ws.Evaluate(@"=MIN(D3:D45)").CastTo<int>();
            Assert.AreEqual(0, value);

            value = ws.Evaluate(@"=MIN(G3:G45)").CastTo<int>();
            Assert.AreEqual(2, value);

            value = ws.Evaluate(@"=MIN(G:G)").CastTo<int>();
            Assert.AreEqual(2, value);

            value = workbook.Evaluate(@"=MIN(Data!G:G)").CastTo<int>();
            Assert.AreEqual(2, value);
        }

        [Test]
        public void StDev()
        {
            var ws = workbook.Worksheets.First();
            double value;
            Assert.That(() => ws.Evaluate(@"=STDEV(D3:D45)"), Throws.TypeOf<ApplicationException>());

            value = ws.Evaluate(@"=STDEV(H3:H45)").CastTo<double>();
            Assert.AreEqual(47.34511769, value, tolerance);

            value = ws.Evaluate(@"=STDEV(H:H)").CastTo<double>();
            Assert.AreEqual(47.34511769, value, tolerance);

            value = workbook.Evaluate(@"=STDEV(Data!H:H)").CastTo<double>();
            Assert.AreEqual(47.34511769, value, tolerance);
        }

        [Test]
        public void StDevP()
        {
            var ws = workbook.Worksheets.First();
            double value;
            Assert.That(() => ws.Evaluate(@"=STDEVP(D3:D45)"), Throws.InvalidOperationException);

            value = ws.Evaluate(@"=STDEVP(H3:H45)").CastTo<double>();
            Assert.AreEqual(46.79135458, value, tolerance);

            value = ws.Evaluate(@"=STDEVP(H:H)").CastTo<double>();
            Assert.AreEqual(46.79135458, value, tolerance);

            value = workbook.Evaluate(@"=STDEVP(Data!H:H)").CastTo<double>();
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
                int initialCount = (ws as XLWorksheet).Internals.CellsCollection.Count;

                var actualResult = (double)ws.Evaluate(formulaA1);
                int cellsCount = (ws as XLWorksheet).Internals.CellsCollection.Count;

                Assert.AreEqual(expectedResult, actualResult, tolerance);
                Assert.AreEqual(initialCount, cellsCount);
            }
        }

        [Test]
        public void Var()
        {
            var ws = workbook.Worksheets.First();
            double value;
            Assert.That(() => ws.Evaluate(@"=VAR(D3:D45)"), Throws.InvalidOperationException);

            value = ws.Evaluate(@"=VAR(H3:H45)").CastTo<double>();
            Assert.AreEqual(2241.560169, value, tolerance);

            value = ws.Evaluate(@"=VAR(H:H)").CastTo<double>();
            Assert.AreEqual(2241.560169, value, tolerance);

            value = workbook.Evaluate(@"=VAR(Data!H:H)").CastTo<double>();
            Assert.AreEqual(2241.560169, value, tolerance);
        }

        [Test]
        public void VarP()
        {
            var ws = workbook.Worksheets.First();
            double value;
            Assert.That(() => ws.Evaluate(@"=VARP(D3:D45)"), Throws.InvalidOperationException);

            value = ws.Evaluate(@"=VARP(H3:H45)").CastTo<double>();
            Assert.AreEqual(2189.430863, value, tolerance);

            value = ws.Evaluate(@"=VARP(H:H)").CastTo<double>();
            Assert.AreEqual(2189.430863, value, tolerance);

            value = workbook.Evaluate(@"=VARP(Data!H:H)").CastTo<double>();
            Assert.AreEqual(2189.430863, value, tolerance);
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
