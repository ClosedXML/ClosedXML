// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    [SetCulture("en-US")]
    public class LookupTests
    {
        private IXLWorksheet ws;

        #region Setup and teardown

        [OneTimeTearDown]
        public void Dispose()
        {
            ws.Workbook.Dispose();
        }

        [SetUp]
        public void Init()
        {
            ws = SetupWorkbook();
        }

        private IXLWorksheet SetupWorkbook()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Data");
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
            ws.FirstCell()
                .CellBelow()
                .CellRight()
                .InsertTable(data);

            return ws;
        }

        #endregion Setup and teardown

        [Test]
        public void Column()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Data");
            wb.AddWorksheet("Other");

            Assert.Multiple(() =>
            {
                // If no argument, function uses the address of the cell that contains the formula
                Assert.That(ws.Cell("D1").SetFormulaA1("COLUMN()").Value, Is.EqualTo(4));

                // With a reference, it returns the column number
                Assert.That(ws.Cell("A1").SetFormulaA1("COLUMN(Z14)").Value, Is.EqualTo(26));

                // If a single column is used, return the column number 
                Assert.That(ws.Cell("A2").SetFormulaA1("COLUMN(C:C)").Value, Is.EqualTo(3));

                // Return a horizontal array for multiple columns. Use SUM to verify content of an array since ROWS/COLUMNS don't work yet.
                Assert.That(ws.Cell("A3").SetFormulaA1("SUM(COLUMN(C:D))").Value, Is.EqualTo(3 + 4));
                Assert.That(ws.Cell("A3").SetFormulaA1("SUM(COLUMN(E1:G10))").Value, Is.EqualTo(5 + 6 + 7));

                // Not contiguous range (multiple areas) returns #REF!
                Assert.That(ws.Cell("A4").SetFormulaA1("COLUMN((D5:G10,I8:K12))").Value, Is.EqualTo(XLError.CellReference));

                // Invalid references return #REF!
                Assert.That(ws.Cell("A5").SetFormulaA1("COLUMN(NonExistent!F10)").Value, Is.EqualTo(XLError.CellReference));

                // Return column number even for different worksheet
                Assert.That(ws.Cell("A6").SetFormulaA1("COLUMN(Other!E7)").Value, Is.EqualTo(5));

                // Unexpected types return error
                Assert.That(ws.Cell("A8").SetFormulaA1("COLUMN(TRUE)").Value, Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(ws.Cell("A7").SetFormulaA1("COLUMN(5)").Value, Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(ws.Cell("A8").SetFormulaA1("COLUMN(\"C5\")").Value, Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(ws.Cell("A9").SetFormulaA1("COLUMN(#DIV/0!)").Value, Is.EqualTo(XLError.DivisionByZero));
                Assert.That(ws.Cell("A10").SetFormulaA1("COLUMN(\"C5\")").Value, Is.EqualTo(XLError.IncompatibleValue));
            });
        }

        [Test]
        public void Columns_Blank_ReturnsValueError()
        {
            Assert.That(XLWorkbook.EvaluateExpr("COLUMNS(IF(TRUE,,))"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [TestCase("0")]
        [TestCase("1")]
        [TestCase("99")]
        [TestCase("-10")]
        [TestCase("TRUE")]
        [TestCase("FALSE")]
        [TestCase("\"\"")]
        [TestCase("\"A\"")]
        [TestCase("\"Hello World\"")]
        public void Columns_ScalarValues_ReturnsOne(string value)
        {
            Assert.That(XLWorkbook.EvaluateExpr($"COLUMNS({value})"), Is.EqualTo(1));
        }

        [Test]
        public void Columns_Error_ReturnsError()
        {
            Assert.That(XLWorkbook.EvaluateExpr("COLUMNS(#DIV/0!)"), Is.EqualTo(XLError.DivisionByZero));
        }

        [TestCase("{1}", 1)]
        [TestCase("{1;2;3}", 1)]
        [TestCase("{1,2,3,4;5,6,7,8}", 4)]
        [TestCase("{TRUE,\"Z\";#DIV/0!,4}", 2)]
        public void Columns_Arrays_ReturnsNumberOfColumns(string array, int expectedColumnCount)
        {
            Assert.That(XLWorkbook.EvaluateExpr($"COLUMNS({array})"), Is.EqualTo(expectedColumnCount));
        }

        [TestCase("A1", 1)]
        [TestCase("A1:A6", 1)]
        [TestCase("B2:D6", 3)]
        [TestCase("E7:AA14", 23)]
        public void Columns_References_ReturnsNumberOfColumns(string range, int expectedColumnCount)
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            Assert.That(sheet.Evaluate($"COLUMNS({range})"), Is.EqualTo(expectedColumnCount));
        }

        [Test]
        public void Columns_NonContiguousReferences_ReturnsReferenceError()
        {
            // Spec says #NULL!, but Excel says #REF!
            Assert.That(XLWorkbook.EvaluateExpr("COLUMNS((A1,C3))"), Is.EqualTo(XLError.CellReference));
        }

        [Test]
        public void Hlookup()
        {
            // Since HLOOKUP requires values to be sorted, we can't use created data.
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            sheet.Cell("B2").InsertData(new[]
            {
                new object[] { 1, 3, 5, 10 },
                new object[] { "A", "B", "C", "D" },
            });

            // Range lookup false = exact match
            var value = sheet.Evaluate(@"HLOOKUP(3,B2:E3,2,FALSE)");
            Assert.That(value, Is.EqualTo("B"));

            // Text values are looked up case insensitive.
            value = sheet.Evaluate(@"HLOOKUP(""c"",B3:E3,1,FALSE)");
            Assert.Multiple(() =>
            {
                Assert.That(value, Is.EqualTo("C"));

                // Value not present in the range for exact search
                // Empty string is not same as blank.
                Assert.That(ws.Evaluate(@"HLOOKUP("""",A2:E2,1,FALSE)"), Is.EqualTo(XLError.NoValueAvailable));
                Assert.That(ws.Evaluate(@"HLOOKUP(50,B2:E3,1,FALSE)"), Is.EqualTo(XLError.NoValueAvailable));

                // Value in approximate search that is lower than first element
                Assert.That(ws.Evaluate(@"HLOOKUP(-10,B2:E3,2,TRUE)"), Is.EqualTo(XLError.NoValueAvailable));
            });
        }

        [Test]
        public void Hlookup_UnexpectedArguments()
        {
            Assert.Multiple(() =>
            {
                // Lookup value can't be an error
                Assert.That(XLWorkbook.EvaluateExpr(@"HLOOKUP(#DIV/0!,{1,2},1)"), Is.EqualTo(XLError.DivisionByZero));

                // Text value can't be over 255 chars
                Assert.That(XLWorkbook.EvaluateExpr($"HLOOKUP(\"{new string('A', 256)}\",{{\"A\"}},1)"), Is.EqualTo(XLError.IncompatibleValue));

                // Range can only be array or a reference. If other type, it returns the error #N/A
                Assert.That(XLWorkbook.EvaluateExpr(@"HLOOKUP(""value"",1,1)"), Is.EqualTo(XLError.NoValueAvailable));
                Assert.That(XLWorkbook.EvaluateExpr(@"HLOOKUP(""value"",TRUE,1)"), Is.EqualTo(XLError.NoValueAvailable));

                // If range is a non-contiguous range, #N/A
                Assert.That(ws.Evaluate(@"HLOOKUP(""Units"",(B2:I5,B6:I10),1)"), Is.EqualTo(XLError.NoValueAvailable));

                // The row index number must be at most the same as height of the range. It is 5 here, but range is 4 cell high.
                Assert.That(ws.Evaluate(@"HLOOKUP(""value"",B2:I5,5,FALSE)"), Is.EqualTo(XLError.CellReference));

                // The row index number must be at least 1. It is 0 here.
                Assert.That(XLWorkbook.EvaluateExpr(@"HLOOKUP(1,{1,2},0,FALSE)"), Is.EqualTo(XLError.IncompatibleValue));
            });
        }

        [Test]
        public void Hlookup_truncates_row_index_number_parameter()
        {
            // If row index number is not a whole number, it is truncated, so here 1.9 is truncated to 1
            Assert.That(ws.Evaluate(@"HLOOKUP(7,{5,7,9},1.9)"), Is.EqualTo(7));
        }

        [Test]
        public void Hlookup_converts_blank_lookup_value_to_number_zero()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A1").InsertData(new[]
            {
                new object[] { -1, 0, 1 },
                new object[] { "-one", "zero", "one"},
            });

            var actual = worksheet.Evaluate("HLOOKUP(IF(TRUE,,),A1:C2,2)");

            Assert.That(actual, Is.EqualTo("zero"));
        }

        [Test]
        public void Hlookup_approximate_search_omits_values_with_different_type()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A1").Value = "0";
            worksheet.Cell("B1").Value = "1";
            worksheet.Cell("C1").Value = 1;
            worksheet.Cell("D1").Value = "0";
            worksheet.Cell("E1").Value = "text";
            worksheet.Cell("F1").Value = Blank.Value;
            worksheet.Cell("G1").Value = 2;
            worksheet.Cell("A2").InsertData(Enumerable.Range(1, 7).Select(x => $"Column {x}"), true);

            var actual = worksheet.Evaluate("HLOOKUP(1.9,A1:G2,2,TRUE)");
            Assert.That(actual, Is.EqualTo("Column 3"));
        }

        [Test]
        public void Hlookup_with_range_containing_only_cells_with_different_type_returns_NA_error()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            sheet.Cell("A1").Value = "text";
            Assert.That(sheet.Evaluate("HLOOKUP(1,A1,1,TRUE)"), Is.EqualTo(XLError.NoValueAvailable));
        }

        [Test]
        public void Hlookup_approximate_search_returns_last_column_for_multiple_equal_values()
        {
            var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            sheet.Cell("A1").InsertData(new object[]
            {
                new object[] { 1, 3, 3, 3, 3, 3, 3, 9 },
                new object[] { "A", "B", "C", "D", "E", "F", "G", "H" },
            });

            // If there is a section of values with same value, return the value at the highest column
            var actual = sheet.Evaluate("HLOOKUP(3, A1:H2, 2, TRUE)");
            Assert.That(actual, Is.EqualTo("G"));

            // If the last value is in the highest column, just return value outright
            actual = sheet.Evaluate("HLOOKUP(3, B1:G2, 2, TRUE)");
            Assert.That(actual, Is.EqualTo("G"));
        }

        [Test]
        public void Hyperlink()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();

            var cell = sheet.Cell("B3");
            cell.FormulaA1 = "HYPERLINK(\"http://github.com/ClosedXML/ClosedXML\")";
            Assert.That(cell.Value, Is.EqualTo("http://github.com/ClosedXML/ClosedXML"));
            Assert.That(cell.HasHyperlink, Is.False);

            cell = sheet.Cell("B4");
            cell.FormulaA1 = "HYPERLINK(\"mailto:jsmith@github.com\", \"jsmith@github.com\")";
            Assert.That(cell.Value, Is.EqualTo("jsmith@github.com"));
            Assert.That(cell.HasHyperlink, Is.False);

            cell = sheet.Cell("B5");
            cell.FormulaA1 = "HYPERLINK(\"[Test.xlsx]Sheet1!A5\", \"Cell A5\")";
            Assert.That(cell.Value, Is.EqualTo("Cell A5"));
            Assert.That(cell.HasHyperlink, Is.False);
        }

        [Test]
        public void Index_reference()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            sheet.Cell("B2").Value = "B2";
            sheet.Cell("B4").Value = "B4";
            sheet.Cell("B5").Value = "B5";
            sheet.Cell("E2").Value = "E2";
            sheet.Cell("E4").Value = "E4";

            // A single cell
            AssertIndex("INDEX(B2:J12, 3, 4)", 1, 1, "E4");

            // Row number is omitted, so take all rows from the range. The result is a column E2:E12
            AssertIndex("INDEX(B2:J12, 0, 4)", 11, 1, "E2");
            AssertIndex("INDEX(B2:J12, , 4)", 11, 1, "E2");

            // Column number is omitted, so take all column from the range. The result is a column B4:J4
            AssertIndex("INDEX(B2:J12, 3, 0)", 1, 9, "B4");
            AssertIndex("INDEX(B2:J12, 3, )", 1, 9, "B4");

            // The range is a row and there is only one parameter. Take the index from the row.
            AssertIndex("INDEX(B2:I2, 4)", 1, 1, "E2");

            // The range is a column and there is only one parameter. Take the index from the column.
            AssertIndex("INDEX(B2:B12, 4)", 1, 1, "B5");

            // Take whole range.
            AssertIndex("INDEX(B2:J12, 0, 0)", 11, 9, "B2");

            // Select second area from multi-area reference
            AssertIndex("INDEX((H4:J10, B2:J12, A1), 1, 1, 2)", 1, 1, "B2");
            return;

            void AssertIndex(string formula, int rows, int cols, XLCellValue value)
            {
                Assert.Multiple(() =>
                {
                    Assert.That(sheet.Evaluate($"INDEX({formula},1,1)"), Is.EqualTo(value));
                    Assert.That(sheet.Evaluate($"ROWS({formula})"), Is.EqualTo(rows));
                    Assert.That(sheet.Evaluate($"COLUMNS({formula})"), Is.EqualTo(cols));
                    Assert.That(sheet.Evaluate($"ISREF({formula})"), Is.EqualTo(true));
                });
            }
        }

        [Test]
        public void Index_reference_errors()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();

            Assert.Multiple(() =>
            {
                // Row bounds
                Assert.That(sheet.Evaluate("INDEX(A1, -1, 1)"), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(sheet.Evaluate("INDEX(B3:C5, 4, 1)"), Is.EqualTo(XLError.CellReference));

                // Column bounds
                Assert.That(sheet.Evaluate("INDEX(A1, 1, -1)"), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(sheet.Evaluate("INDEX(B3:C5, 1, 3)"), Is.EqualTo(XLError.CellReference));

                // Area bounds
                Assert.That(sheet.Evaluate("INDEX((A1, B1, C1), 1, 1, 0)"), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(sheet.Evaluate("INDEX((A1, B1, C1),1, 1, 4)"), Is.EqualTo(XLError.CellReference));
            });
        }

        [Test]
        public void Index_array()
        {
            // A single element
            AssertIndex("INDEX({1,2,3;4,5,6}, 2, 3)", 1, 1, 6);

            // Row number is omitted, so take all rows from the array at third column. The result is a column {3;6}
            AssertIndex("INDEX({1,2,3;4,5,6}, 0, 3)", 2, 1, 3);
            AssertIndex("INDEX({1,2,3;4,5,6}, , 3)", 2, 1, 3);

            // Column number is omitted, so take all columns from the array at second row. The result is a row {4,5,6}
            AssertIndex("INDEX({1,2,3;4,5,6}, 2, 0)", 1, 3, 4);
            AssertIndex("INDEX({1,2,3;4,5,6}, 2, )", 1, 3, 4);

            // The array is a row and there is only one parameter. Take the index from the row.
            AssertIndex("INDEX({1,2,3,4,5,6,7}, 5)", 1, 1, 5);

            // The array is a column and there is only one parameter. Take the index from the column.
            AssertIndex("INDEX({1;2;3;4;5;6;7}, 6)", 1, 1, 6);

            // Take whole range.
            AssertIndex("INDEX({1,2,3;4,5,6}, 0, 0)", 2, 3, 1);

            return;

            void AssertIndex(string formula, int rows, int cols, XLCellValue value)
            {
                Assert.Multiple(() =>
                {
                    Assert.That(XLWorkbook.EvaluateExpr(formula), Is.EqualTo(value));
                    Assert.That(XLWorkbook.EvaluateExpr($"ROWS({formula})"), Is.EqualTo(rows));
                    Assert.That(XLWorkbook.EvaluateExpr($"COLUMNS({formula})"), Is.EqualTo(cols));
                    Assert.That(XLWorkbook.EvaluateExpr($"ISREF({formula})"), Is.EqualTo(false));
                });
            }
        }

        [Test]
        public void Index_array_errors()
        {
            Assert.Multiple(() =>
            {
                // Row bounds
                Assert.That(XLWorkbook.EvaluateExpr("INDEX({1}, -1, 1)"), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(XLWorkbook.EvaluateExpr("INDEX({1,2;3,4;5,6}, 4, 1)"), Is.EqualTo(XLError.CellReference));

                // Column bounds
                Assert.That(XLWorkbook.EvaluateExpr("INDEX({1}, 1, -1)"), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(XLWorkbook.EvaluateExpr("INDEX({1,2;3,4;5,6}, 1, 3)"), Is.EqualTo(XLError.CellReference));

                // Area bounds
                Assert.That(XLWorkbook.EvaluateExpr("INDEX({1}, 1, 1, 0)"), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(XLWorkbook.EvaluateExpr("INDEX({1}, 1, 1, 2)"), Is.EqualTo(XLError.CellReference));
            });
        }

        [Test]
        public void Index_scalar()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr("INDEX(\"Text\", 1, 1)"), Is.EqualTo("Text"));
                Assert.That(XLWorkbook.EvaluateExpr("INDEX(\"Text\", 0, 0)"), Is.EqualTo("Text"));
                Assert.That(XLWorkbook.EvaluateExpr("TYPE(INDEX(\"Text\", 1, 1))"), Is.EqualTo(2));
                Assert.That(XLWorkbook.EvaluateExpr("INDEX(IF(TRUE,), 1, 1)"), Is.EqualTo(XLError.IncompatibleValue));

                Assert.That(XLWorkbook.EvaluateExpr("INDEX(\"Text\", -1, 1)"), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(XLWorkbook.EvaluateExpr("INDEX(\"Text\", 2, 1)"), Is.EqualTo(XLError.CellReference));
                Assert.That(XLWorkbook.EvaluateExpr("INDEX(\"Text\", 1, -1)"), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(XLWorkbook.EvaluateExpr("INDEX(\"Text\", 1, 2)"), Is.EqualTo(XLError.CellReference));
                Assert.That(XLWorkbook.EvaluateExpr("INDEX(\"Text\", 1, 1, 0)"), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(XLWorkbook.EvaluateExpr("INDEX(\"Text\", 1, 1, 2)"), Is.EqualTo(XLError.CellReference));
            });
        }

        [TestCase(@"MATCH(""Rep"", B2:I2, 0)", 4)]
        [TestCase(@"MATCH(""Rep"", A2:Z2, 0)", 5)]
        [TestCase(@"MATCH(""REP"", B2:I2, 0)", 4)]
        [TestCase(@"MATCH(95, B3:I3, 0)", 6)]
        [TestCase(@"MATCH(DATE(2015,1,6), B3:I3, 0)", 2)]
        [TestCase(@"MATCH(1.99, 3:3, 0)", 8)]
        [TestCase(@"MATCH(43, B:B, 0)", 45)]
        [TestCase(@"MATCH(""cENtraL"", D3:D45, 0)", 2)]
        [TestCase(@"MATCH(4.99, H:H, 0)", 5)]
        [TestCase(@"MATCH(""Rapture"", B2:I2, 1)", 2)]
        [TestCase(@"MATCH(22.5, B3:B45, 1)", 22)]
        [TestCase(@"MATCH(""Rep"", B2:I2)", 4)]
        [TestCase(@"MATCH(""Rep"", B2:I2, 1)", 4)]
        [TestCase(@"MATCH(E2, B2:I2, 0)", 4)]
        [TestCase(@"MATCH(40, G3:G6, -1)", 2)]
        [TestCase(@"MATCH(""Rep"", B2:I5)", XLError.NoValueAvailable)]
        [TestCase(@"MATCH(""Dummy"", B2:I2, 0)", XLError.NoValueAvailable)]
        [TestCase(@"MATCH(4.5,B3:B45,-1)", XLError.NoValueAvailable)]
        public void Match_demo_sheet(string formula, object result)
        {
            var actual = ws.Evaluate(formula);
            Assert.That(actual, Is.EqualTo(result));
        }

        [Test]
        public void Match_examples()
        {
            Assert.Multiple(() =>
            {
                // Examples from specification
                Assert.That(XLWorkbook.EvaluateExpr("MATCH(39,{25,38,40,41},1)"), Is.EqualTo(2));
                Assert.That(XLWorkbook.EvaluateExpr("MATCH(41,{25,38,40,41},0)"), Is.EqualTo(4));
            });

            // Example from office website
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            sheet.Cell("A1").InsertData(new object[]
            {
                ("Product", "Count"),
                ("Bananas", 25),
                ("Oranges", 38),
                ("Apples", 40),
                ("Pears", 41),
            });

            Assert.Multiple(() =>
            {
                Assert.That(sheet.Evaluate("MATCH(39,B2:B5,1)"), Is.EqualTo(2));
                Assert.That(sheet.Evaluate("MATCH(41,B2:B5,0)"), Is.EqualTo(4));
                Assert.That(sheet.Evaluate("MATCH(40,B2:B5,-1)"), Is.EqualTo(XLError.NoValueAvailable));
            });
        }

        [TestCase("MATCH(5, {10,5,4,5,5,5,5,5}, -1)", 2)] // Doesn't use bisection, otherwise it would pick later position
        [TestCase("MATCH(5, {10,4,5}, -1)", 1)] // Because 4 is less than the target, search stops. Values should be descending.
        [TestCase("MATCH(5, {\"5\",10,\"4\",FALSE,TRUE,#DIV/0!,5,3}, -1)", 7)] // Non-target values are ignored
        [TestCase("MATCH(6, {\"4\",10,\"4\",FALSE,TRUE,#DIV/0!,5,3}, -1)", 2)] // Returned position is of the correct type, not just before less than target.
        [TestCase("MATCH(5, {\"5\"}, -1)", XLError.NoValueAvailable)] // String values are not converted to numbers
        [TestCase("MATCH(5, {4}, -1)", XLError.NoValueAvailable)]
        [TestCase("MATCH(5, {10}, -1)", 1)]
        [TestCase("MATCH(5, {TRUE}, -1)", XLError.NoValueAvailable)]
        [TestCase("MATCH(\"c\", {\"E\",4,\"D\",\"B\"}, -1)", 3)]
        [TestCase("MATCH(FALSE, {TRUE,TRUE,\"FALSE\",0,FALSE,FALSE}, -1)", 5)]
        public void Match_from_descending(string formula, object result)
        {
            var actual = XLWorkbook.EvaluateExpr(formula);
            Assert.That(actual, Is.EqualTo(result));
        }

        [TestCase("MATCH(35,{25,38,24,35,70},0)", 4)] // Finds value even in unsorted
        [TestCase("MATCH(35,{\"35\",38,24,35,70},0)", 4)] // String values are not converted, must match type
        [TestCase("MATCH(1,{5},0)", XLError.NoValueAvailable)] // Nothing found
        [TestCase("MATCH(\"35\",{35,38,24,\"35\",70},0)", 4)] // String target is not converted, must match type
        [TestCase("MATCH(\"c*\",{\"a\",\"cd\"},0)", 2)] // Consider string targets wildcards
        [TestCase("MATCH(TRUE, {0,\"TRUE\",FALSE,TRUE,1},0)", 4)]
        public void Match_from_unsorted(string formula, object result)
        {
            var actual = XLWorkbook.EvaluateExpr(formula);
            Assert.That(actual, Is.EqualTo(result));
        }

        [TestCase("MATCH(39,{25,38,38,38,40,41},1)", 4)] // When there is a sequence of target values, return last one
        [TestCase("MATCH(20,{25,38,40},1)", XLError.NoValueAvailable)] // Nothing found, even smallest value is greater than target
        [TestCase("MATCH(25,{20,TRUE,FALSE,38,40},1)", 1)] // If found value is <= target, return position of value, not subsequent types that are ignored
        [TestCase("MATCH(8, {FALSE;FALSE}, 1)", XLError.NoValueAvailable)] // Not even one value of target type
        [TestCase("MATCH(5, {1,2,3}, 1)", 3)] // If target value is greater than the last element of same type, return the position of the last element
        public void Match_from_ascending(string formula, object result)
        {
            var actual = XLWorkbook.EvaluateExpr(formula);
            Assert.That(actual, Is.EqualTo(result));
        }

        [TestCase("MATCH(17, {14;5;3;5;11;12;11;13;13;4})", 10)]
        [TestCase("MATCH(12, {5;15;18;18;11;1;15;17})", 1)]
        [TestCase("MATCH(4, {10,3,FALSE, FALSE,FALSE})", XLError.NoValueAvailable)]
        [TestCase("MATCH(8, {14;0;17;FALSE;8})", XLError.NoValueAvailable)]
        public void Match_from_ascending_matches_excel(string formula, object result)
        {
            // The bisection algorithm should match Excel. That is checked by supplying
            // non-ascending data and checking the result against Excel result. Use random
            // generator to generate formulas + compare with Excel when modifying the algorithm.
            var actual = XLWorkbook.EvaluateExpr(formula);
            Assert.That(actual, Is.EqualTo(result));
        }

        [TestCase("MATCH(#DIV/0!,{1,2,3},1)", XLError.DivisionByZero)] // Scalar argument is error -> propagate
        [TestCase("MATCH(IF(TRUE,),{1,2,3},1)", XLError.NoValueAvailable)] // Return not found for blank value
        [TestCase("MATCH(1,{1,2;3,4},1)", XLError.NoValueAvailable)] // Must be either row or column, the array is 2x2
        [TestCase("MATCH(1,{3,2,1},-2)", 3)] // Match type can be negative for match type -1
        [TestCase("MATCH(1,{1,2,3}, 2)", 1)] // Match type can be positive for match type 1
        [TestCase("MATCH(2,{1;2;3}, 2)", 2)] // Match returns position from start both in row or column
        [TestCase("MATCH(2,{1,2,3}, 2)", 2)] // Match returns position from start both in row or column
        [TestCase("MATCH(3,{1,2,3,4,5})", 3)] // Default match type is 1 (ascending bisection)
        [TestCase("MATCH(3,3)", XLError.NoValueAvailable)] // Scalar values are not converted to 1x1 array
        public void Match_edge_conditions(string formula, object result)
        {
            var actual = XLWorkbook.EvaluateExpr(formula);
            Assert.That(actual, Is.EqualTo(result));
        }

        [Test]
        public void Match_accepts_single_cell_as_values()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            sheet.Cell("A1").Value = 5;
            Assert.That(sheet.Evaluate("MATCH(5, A1)"), Is.EqualTo(1));
        }

        [Test]
        public void Row()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Data");
            wb.AddWorksheet("Other");

            Assert.Multiple(() =>
            {
                // If no argument, function uses the address of the cell that contains the formula
                Assert.That(ws.Cell("M60").SetFormulaA1("ROW()").Value, Is.EqualTo(60));

                // With a reference, it returns the row number
                Assert.That(ws.Cell("A1").SetFormulaA1("ROW(C12)").Value, Is.EqualTo(12));

                // If a full row reference to a single row is used, return the row number 
                Assert.That(ws.Cell("A2").SetFormulaA1("ROW(40:40)").Value, Is.EqualTo(40));

                // Return a vertical array for multiple rows. Use SUM to verify content of an array since ROWS/COLUMNS don't work yet.
                Assert.That(ws.Cell("A3").SetFormulaA1("SUM(ROW(4:7))").Value, Is.EqualTo(4 + 5 + 6 + 7));
                Assert.That(ws.Cell("A4").SetFormulaA1("SUM(ROW(C2:Z4))").Value, Is.EqualTo(2 + 3 + 4));

                // Not contiguous range (multiple areas) returns #REF!
                Assert.That(ws.Cell("A5").SetFormulaA1("ROW((D5:G10,I8:K12))").Value, Is.EqualTo(XLError.CellReference));

                // Invalid references return #REF!
                Assert.That(ws.Cell("A6").SetFormulaA1("ROW(NonExistent!F10)").Value, Is.EqualTo(XLError.CellReference));

                // Return row number even for different worksheet
                Assert.That(ws.Cell("A7").SetFormulaA1("ROW(Other!E14)").Value, Is.EqualTo(14));

                // Unexpected types return error
                Assert.That(ws.Cell("A8").SetFormulaA1("ROW(IF(TRUE,TRUE))").Value, Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(ws.Cell("A9").SetFormulaA1("ROW(IF(TRUE,5))").Value, Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(ws.Cell("A10").SetFormulaA1("ROW(IF(TRUE,\"G15\"))").Value, Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(ws.Cell("A11").SetFormulaA1("ROW(#DIV/0!)").Value, Is.EqualTo(XLError.DivisionByZero));
            });

            // Properly works even in array formulas, where border between references and arrays blurs.
            ws.Range("A12:A13").FormulaArrayA1 = "ROW(2:3)";
            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A12").Value, Is.EqualTo(2));
                Assert.That(ws.Cell("A13").Value, Is.EqualTo(3));
            });
        }

        [Test]
        public void Rows_Blank_ReturnsValueError()
        {
            Assert.That(XLWorkbook.EvaluateExpr("ROWS(IF(TRUE,,))"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [TestCase("0")]
        [TestCase("1")]
        [TestCase("99")]
        [TestCase("-10")]
        [TestCase("TRUE")]
        [TestCase("FALSE")]
        [TestCase("\"\"")]
        [TestCase("\"A\"")]
        [TestCase("\"Hello World\"")]
        public void Rows_ScalarValues_ReturnsOne(string value)
        {
            Assert.That(XLWorkbook.EvaluateExpr($"ROWS({value})"), Is.EqualTo(1));
        }

        [Test]
        public void Rows_Error_ReturnsError()
        {
            Assert.That(XLWorkbook.EvaluateExpr("ROWS(#DIV/0!)"), Is.EqualTo(XLError.DivisionByZero));
        }

        [TestCase("{1}", 1)]
        [TestCase("{1;2;3}", 3)]
        [TestCase("{1,2,3,4;5,6,7,8;9,10,11,12}", 3)]
        [TestCase("{TRUE;#DIV/0!}", 2)]
        public void Rows_Arrays_ReturnsNumberOfRows(string array, int expectedColumnCount)
        {
            Assert.That(XLWorkbook.EvaluateExpr($"ROWS({array})"), Is.EqualTo(expectedColumnCount));
        }

        [TestCase("C3", 1)]
        [TestCase("B3:E12", 10)]
        [TestCase("AA21:AC400", 380)]
        public void Rows_References_ReturnsNumberOfColumns(string range, int expectedColumnCount)
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            Assert.That(sheet.Evaluate($"ROWS({range})"), Is.EqualTo(expectedColumnCount));
        }

        [Test]
        public void Rows_NonContiguousReferences_ReturnsReferenceError()
        {
            // Spec says #NULL!, but Excel says #REF!
            Assert.That(XLWorkbook.EvaluateExpr("ROWS((A1,C3))"), Is.EqualTo(XLError.CellReference));
        }

        [Test]
        public void Vlookup()
        {
            // Range lookup false = exact match
            var value = ws.Evaluate("=VLOOKUP(3,Data!$B$2:$I$71,3,FALSE)");
            Assert.That(value, Is.EqualTo("Central"));

            value = ws.Evaluate("=VLOOKUP(DATE(2015,5,22),Data!C:I,7,FALSE)");
            Assert.That(value, Is.EqualTo(63.68));

            value = ws.Evaluate(@"=VLOOKUP(""Central"",Data!D:E,2,FALSE)");
            Assert.That(value, Is.EqualTo("Kivell"));

            // Case insensitive lookup
            value = ws.Evaluate(@"=VLOOKUP(""central"",Data!D:E,2,FALSE)");
            Assert.That(value, Is.EqualTo("Kivell"));

            // Range lookup true = approximate match
            value = ws.Evaluate("=VLOOKUP(3,Data!$B$2:$I$71,8,TRUE)");
            Assert.That(value, Is.EqualTo(179.64));

            value = ws.Evaluate("=VLOOKUP(3,Data!$B$2:$I$71,8)");
            Assert.That(value, Is.EqualTo(179.64));

            value = ws.Evaluate("=VLOOKUP(3,Data!$B$2:$I$71,8,)");
            Assert.That(value, Is.EqualTo(179.64));

            value = ws.Evaluate("=VLOOKUP(14.5,Data!$B$2:$I$71,8,TRUE)");
            Assert.That(value, Is.EqualTo(174.65));

            value = ws.Evaluate("=VLOOKUP(50,Data!$B$2:$I$71,8,TRUE)");
            Assert.That(value, Is.EqualTo(139.72));
        }

        [Test]
        public void Vlookup_ElementNotFound_ReturnsNotAvailableError()
        {
            Assert.Multiple(() =>
            {
                // Value not present in the range for exact search
                Assert.That(ws.Evaluate(@"=VLOOKUP("""",Data!$B$2:$I$71,3,FALSE)"), Is.EqualTo(XLError.NoValueAvailable));
                Assert.That(ws.Evaluate(@"=VLOOKUP(50,Data!$B$2:$I$71,3,FALSE)"), Is.EqualTo(XLError.NoValueAvailable));

                // Value in approximate search that is lower than first element
                Assert.That(ws.Evaluate(@"=VLOOKUP(-1,Data!$B$2:$I$71,2,TRUE)"), Is.EqualTo(XLError.NoValueAvailable));
            });
        }

        [Test]
        public void Vlookup_UnexpectedArguments()
        {
            Assert.Multiple(() =>
            {
                // Lookup value can't be an error
                Assert.That(ws.Evaluate("=VLOOKUP(#DIV/0!,B2:I71,1)"), Is.EqualTo(XLError.DivisionByZero));

                // Text value can't be over 255 chars
                Assert.That(ws.Evaluate($"=VLOOKUP(\"{new string('A', 256)}\",B2:I71,1)"), Is.EqualTo(XLError.IncompatibleValue));

                // Range can only be array or a reference. If other type, it returns the error #N/A
                Assert.That(ws.Evaluate("=VLOOKUP(1,1,1)"), Is.EqualTo(XLError.NoValueAvailable));
                Assert.That(ws.Evaluate("=VLOOKUP(1,TRUE,1)"), Is.EqualTo(XLError.NoValueAvailable));

                // If range is a non-contiguous range, #N/A
                Assert.That(ws.Evaluate("=VLOOKUP(1,(B2:I5,B6:I10),1)"), Is.EqualTo(XLError.NoValueAvailable));

                // The column index must be at most the same as width of the range. It is 9 here, but range is 8 cell wide.
                Assert.That(ws.Evaluate("=VLOOKUP(20,B2:I71,9,FALSE)"), Is.EqualTo(XLError.CellReference));
                // The column index must be at least 1. It is 0 here.
                Assert.That(ws.Evaluate("=VLOOKUP(20,B2:I71,0,FALSE)"), Is.EqualTo(XLError.IncompatibleValue));
            });
        }

        [Test]
        public void Vlookup_ColumnIndexParameter_UsesValueSemantic()
        {
            Assert.Multiple(() =>
            {
                // If column index is not a whole number, it is truncated, so here 1.9 is truncated to 1
                Assert.That(ws.Evaluate("=VLOOKUP(14,B2:I71,1.9)"), Is.EqualTo(14.0));

                // Column index is evaluated using a VALUE semantic
                Assert.That(ws.Evaluate("=VLOOKUP(3,B2:I71,\"2 5/2\")"), Is.EqualTo(@"Jardine"));
            });
        }

        [TestCase("\"TRUE\"")]
        [TestCase("1")]
        [TestCase("TRUE")]
        public void Vlookup_FlagParameter_CoercedToBoolean(string flagValue)
        {
            Assert.That(ws.Evaluate($"VLOOKUP(5,B2:I71,1,{flagValue})"), Is.EqualTo(5.0));
        }

        [Test]
        public void Vlookup_BlankLookupValue_BehavesAsZero()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A1").InsertData(Enumerable.Range(-5, 10).Select(x => new object[] { x, $"Row with value {x}" }));

            var actual = worksheet.Evaluate("VLOOKUP(IF(TRUE,,),A1:B10,2)");

            Assert.That(actual, Is.EqualTo("Row with value 0"));
        }

        [Test]
        public void Vlookup_ApproximateSearch_OmitsValuesWithDifferentType()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A1").Value = "0";
            worksheet.Cell("A2").Value = "1";
            worksheet.Cell("A3").Value = 1;
            worksheet.Cell("A4").Value = "0";
            worksheet.Cell("A5").Value = "text";
            worksheet.Cell("A6").Value = Blank.Value;
            worksheet.Cell("A7").Value = 2;
            worksheet.Cell("B1").InsertData(Enumerable.Range(1, 7).Select(x => $"Row {x}"));

            var actual = worksheet.Evaluate("VLOOKUP(1.9,A1:B7,2,TRUE)");
            Assert.That(actual, Is.EqualTo("Row 3"));
        }

        [Test]
        public void Vlookup_OnlyCellsWithDifferentType_ReturnsNotAvailable()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            Assert.That(worksheet.Evaluate("VLOOKUP(1,A1,1,TRUE)"), Is.EqualTo(XLError.NoValueAvailable));
        }

        [Test]
        public void Vlookup_OnlyOneValueSurroundedByIgnoredTypes()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A3").Value = 5;

            Assert.That(worksheet.Evaluate("VLOOKUP(6,A1:A5,1,TRUE)"), Is.EqualTo(5));
        }

        [Test]
        public void Vlookup_ResultAtTheHighestCellWithTrailingDifferentTypeAtTheEnd()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A1").Value = 1;
            worksheet.Cell("A2").Value = 2;
            worksheet.Cell("A3").Value = 3;
            worksheet.Cell("A4").Value = Blank.Value;

            Assert.That(worksheet.Evaluate("VLOOKUP(3,A1:A4,1,TRUE)"), Is.EqualTo(3));
        }

        [Test]
        public void Vlookup_ApproximateSearch_ReturnsLastRowForMultipleEqualValues()
        {
            var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            sheet.Cell("A1").Value = 1;
            sheet.Cell("A2").Value = 3;
            sheet.Cell("A3").Value = 3;
            sheet.Cell("A4").Value = 3;
            sheet.Cell("A5").Value = 3;
            sheet.Cell("A6").Value = 3;
            sheet.Cell("A7").Value = 3;
            sheet.Cell("A8").Value = 9;
            sheet.Cell("B1").InsertData(Enumerable.Range(1, 8));

            // If there is a section of values with same value, return the value at the highest row
            var actual = sheet.Evaluate("VLOOKUP(3, A1:B8, 2, TRUE)");
            Assert.That(actual, Is.EqualTo(7));

            // If the last value is in the highest row, just return value outright
            actual = sheet.Evaluate("VLOOKUP(3, A2:B7, 2, TRUE)");
            Assert.That(actual, Is.EqualTo(7));
        }

        [Test]
        public void Vlookup_CanSearchArrays()
        {
            Assert.That(XLWorkbook.EvaluateExpr("VLOOKUP(4, {1,2; 3,2; 5,3; 7,4}, 2)"), Is.EqualTo(2));
        }
    }
}
