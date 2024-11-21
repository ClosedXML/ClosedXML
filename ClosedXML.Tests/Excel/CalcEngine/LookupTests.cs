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

            // If no argument, function uses the address of the cell that contains the formula
            Assert.AreEqual(4, ws.Cell("D1").SetFormulaA1("COLUMN()").Value);

            // With a reference, it returns the column number
            Assert.AreEqual(26, ws.Cell("A1").SetFormulaA1("COLUMN(Z14)").Value);

            // If a single column is used, return the column number 
            Assert.AreEqual(3, ws.Cell("A2").SetFormulaA1("COLUMN(C:C)").Value);

            // Return a horizontal array for multiple columns. Use SUM to verify content of an array since ROWS/COLUMNS don't work yet.
            Assert.AreEqual(3 + 4, ws.Cell("A3").SetFormulaA1("SUM(COLUMN(C:D))").Value);
            Assert.AreEqual(5 + 6 + 7, ws.Cell("A3").SetFormulaA1("SUM(COLUMN(E1:G10))").Value);

            // Not contiguous range (multiple areas) returns #REF!
            Assert.AreEqual(XLError.CellReference, ws.Cell("A4").SetFormulaA1("COLUMN((D5:G10,I8:K12))").Value);

            // Invalid references return #REF!
            Assert.AreEqual(XLError.CellReference, ws.Cell("A5").SetFormulaA1("COLUMN(NonExistent!F10)").Value);

            // Return column number even for different worksheet
            Assert.AreEqual(5, ws.Cell("A6").SetFormulaA1("COLUMN(Other!E7)").Value);

            // Unexpected types return error
            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A8").SetFormulaA1("COLUMN(TRUE)").Value);
            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A7").SetFormulaA1("COLUMN(5)").Value);
            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A8").SetFormulaA1("COLUMN(\"C5\")").Value);
            Assert.AreEqual(XLError.DivisionByZero, ws.Cell("A9").SetFormulaA1("COLUMN(#DIV/0!)").Value);
            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A10").SetFormulaA1("COLUMN(\"C5\")").Value);
        }

        [Test]
        public void Columns_Blank_ReturnsValueError()
        {
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("COLUMNS(IF(TRUE,,))"));
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
            Assert.AreEqual(1, XLWorkbook.EvaluateExpr($"COLUMNS({value})"));
        }

        [Test]
        public void Columns_Error_ReturnsError()
        {
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("COLUMNS(#DIV/0!)"));
        }

        [TestCase("{1}", 1)]
        [TestCase("{1;2;3}", 1)]
        [TestCase("{1,2,3,4;5,6,7,8}", 4)]
        [TestCase("{TRUE,\"Z\";#DIV/0!,4}", 2)]
        public void Columns_Arrays_ReturnsNumberOfColumns(string array, int expectedColumnCount)
        {
            Assert.AreEqual(expectedColumnCount, XLWorkbook.EvaluateExpr($"COLUMNS({array})"));
        }

        [TestCase("A1", 1)]
        [TestCase("A1:A6", 1)]
        [TestCase("B2:D6", 3)]
        [TestCase("E7:AA14", 23)]
        public void Columns_References_ReturnsNumberOfColumns(string range, int expectedColumnCount)
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            Assert.AreEqual(expectedColumnCount, sheet.Evaluate($"COLUMNS({range})"));
        }

        [Test]
        public void Columns_NonContiguousReferences_ReturnsReferenceError()
        {
            // Spec says #NULL!, but Excel says #REF!
            Assert.AreEqual(XLError.CellReference, XLWorkbook.EvaluateExpr("COLUMNS((A1,C3))"));
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
            Assert.AreEqual("B", value);

            // Text values are looked up case insensitive.
            value = sheet.Evaluate(@"HLOOKUP(""c"",B3:E3,1,FALSE)");
            Assert.AreEqual("C", value);

            // Value not present in the range for exact search
            // Empty string is not same as blank.
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"HLOOKUP("""",A2:E2,1,FALSE)"));
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"HLOOKUP(50,B2:E3,1,FALSE)"));

            // Value in approximate search that is lower than first element
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"HLOOKUP(-10,B2:E3,2,TRUE)"));
        }

        [Test]
        public void Hlookup_UnexpectedArguments()
        {
            // Lookup value can't be an error
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr(@"HLOOKUP(#DIV/0!,{1,2},1)"));

            // Text value can't be over 255 chars
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr($"HLOOKUP(\"{new string('A', 256)}\",{{\"A\"}},1)"));

            // Range can only be array or a reference. If other type, it returns the error #N/A
            Assert.AreEqual(XLError.NoValueAvailable, XLWorkbook.EvaluateExpr(@"HLOOKUP(""value"",1,1)"));
            Assert.AreEqual(XLError.NoValueAvailable, XLWorkbook.EvaluateExpr(@"HLOOKUP(""value"",TRUE,1)"));

            // If range is a non-contiguous range, #N/A
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"HLOOKUP(""Units"",(B2:I5,B6:I10),1)"));

            // The row index number must be at most the same as height of the range. It is 5 here, but range is 4 cell high.
            Assert.AreEqual(XLError.CellReference, ws.Evaluate(@"HLOOKUP(""value"",B2:I5,5,FALSE)"));

            // The row index number must be at least 1. It is 0 here.
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr(@"HLOOKUP(1,{1,2},0,FALSE)"));
        }

        [Test]
        public void Hlookup_truncates_row_index_number_parameter()
        {
            // If row index number is not a whole number, it is truncated, so here 1.9 is truncated to 1
            Assert.AreEqual(7, ws.Evaluate(@"HLOOKUP(7,{5,7,9},1.9)"));
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

            Assert.AreEqual("zero", actual);
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
            Assert.AreEqual("Column 3", actual);
        }

        [Test]
        public void Hlookup_with_range_containing_only_cells_with_different_type_returns_NA_error()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            sheet.Cell("A1").Value = "text";
            Assert.AreEqual(XLError.NoValueAvailable, sheet.Evaluate("HLOOKUP(1,A1,1,TRUE)"));
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
            Assert.AreEqual("G", actual);

            // If the last value is in the highest column, just return value outright
            actual = sheet.Evaluate("HLOOKUP(3, B1:G2, 2, TRUE)");
            Assert.AreEqual("G", actual);
        }

        [Test]
        public void Hyperlink()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();

            var cell = sheet.Cell("B3");
            cell.FormulaA1 = "HYPERLINK(\"http://github.com/ClosedXML/ClosedXML\")";
            Assert.AreEqual("http://github.com/ClosedXML/ClosedXML", cell.Value);
            Assert.False(cell.HasHyperlink);

            cell = sheet.Cell("B4");
            cell.FormulaA1 = "HYPERLINK(\"mailto:jsmith@github.com\", \"jsmith@github.com\")";
            Assert.AreEqual("jsmith@github.com", cell.Value);
            Assert.False(cell.HasHyperlink);

            cell = sheet.Cell("B5");
            cell.FormulaA1 = "HYPERLINK(\"[Test.xlsx]Sheet1!A5\", \"Cell A5\")";
            Assert.AreEqual("Cell A5", cell.Value);
            Assert.False(cell.HasHyperlink);
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
                Assert.AreEqual(value, sheet.Evaluate($"INDEX({formula},1,1)"));
                Assert.AreEqual(rows, sheet.Evaluate($"ROWS({formula})"));
                Assert.AreEqual(cols, sheet.Evaluate($"COLUMNS({formula})"));
                Assert.AreEqual(true, sheet.Evaluate($"ISREF({formula})"));
            }
        }

        [Test]
        public void Index_reference_errors()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();

            // Row bounds
            Assert.AreEqual(XLError.IncompatibleValue, sheet.Evaluate("INDEX(A1, -1, 1)"));
            Assert.AreEqual(XLError.CellReference, sheet.Evaluate("INDEX(B3:C5, 4, 1)"));

            // Column bounds
            Assert.AreEqual(XLError.IncompatibleValue, sheet.Evaluate("INDEX(A1, 1, -1)"));
            Assert.AreEqual(XLError.CellReference, sheet.Evaluate("INDEX(B3:C5, 1, 3)"));

            // Area bounds
            Assert.AreEqual(XLError.IncompatibleValue, sheet.Evaluate("INDEX((A1, B1, C1), 1, 1, 0)"));
            Assert.AreEqual(XLError.CellReference, sheet.Evaluate("INDEX((A1, B1, C1),1, 1, 4)"));
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
                Assert.AreEqual(value, XLWorkbook.EvaluateExpr(formula));
                Assert.AreEqual(rows, XLWorkbook.EvaluateExpr($"ROWS({formula})"));
                Assert.AreEqual(cols, XLWorkbook.EvaluateExpr($"COLUMNS({formula})"));
                Assert.AreEqual(false, XLWorkbook.EvaluateExpr($"ISREF({formula})"));
            }
        }

        [Test]
        public void Index_array_errors()
        {
            // Row bounds
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("INDEX({1}, -1, 1)"));
            Assert.AreEqual(XLError.CellReference, XLWorkbook.EvaluateExpr("INDEX({1,2;3,4;5,6}, 4, 1)"));

            // Column bounds
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("INDEX({1}, 1, -1)"));
            Assert.AreEqual(XLError.CellReference, XLWorkbook.EvaluateExpr("INDEX({1,2;3,4;5,6}, 1, 3)"));

            // Area bounds
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("INDEX({1}, 1, 1, 0)"));
            Assert.AreEqual(XLError.CellReference, XLWorkbook.EvaluateExpr("INDEX({1}, 1, 1, 2)"));
        }

        [Test]
        public void Index_scalar()
        {
            Assert.AreEqual("Text", XLWorkbook.EvaluateExpr("INDEX(\"Text\", 1, 1)"));
            Assert.AreEqual("Text", XLWorkbook.EvaluateExpr("INDEX(\"Text\", 0, 0)"));
            Assert.AreEqual(2, XLWorkbook.EvaluateExpr("TYPE(INDEX(\"Text\", 1, 1))"));
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("INDEX(IF(TRUE,), 1, 1)"));

            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("INDEX(\"Text\", -1, 1)"));
            Assert.AreEqual(XLError.CellReference, XLWorkbook.EvaluateExpr("INDEX(\"Text\", 2, 1)"));
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("INDEX(\"Text\", 1, -1)"));
            Assert.AreEqual(XLError.CellReference, XLWorkbook.EvaluateExpr("INDEX(\"Text\", 1, 2)"));
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("INDEX(\"Text\", 1, 1, 0)"));
            Assert.AreEqual(XLError.CellReference, XLWorkbook.EvaluateExpr("INDEX(\"Text\", 1, 1, 2)"));
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
        [TestCase(@"MATCH(40, G3:G6, -1)", 2)]
        [TestCase(@"MATCH(""Rep"", B2:I5)", XLError.NoValueAvailable)]
        [TestCase(@"MATCH(""Dummy"", B2:I2, 0)", XLError.NoValueAvailable)]
        [TestCase(@"MATCH(4.5,B3:B45,-1)", XLError.NoValueAvailable)]
        public void Match_demo_sheet(string formula, object result)
        {
            var actual = ws.Evaluate(formula);
            Assert.AreEqual(result, actual);
        }

        [Test]
        public void Match_examples()
        {
            // Examples from specification
            Assert.AreEqual(2, XLWorkbook.EvaluateExpr("MATCH(39,{25,38,40,41},1)"));
            Assert.AreEqual(4, XLWorkbook.EvaluateExpr("MATCH(41,{25,38,40,41},0)"));

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

            Assert.AreEqual(2, sheet.Evaluate("MATCH(39,B2:B5,1)"));
            Assert.AreEqual(4, sheet.Evaluate("MATCH(41,B2:B5,0)"));
            Assert.AreEqual(XLError.NoValueAvailable, sheet.Evaluate("MATCH(40,B2:B5,-1)"));
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
            Assert.AreEqual(result, actual);
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
            Assert.AreEqual(result, actual);
        }

        [TestCase("MATCH(39,{25,38,38,38,40,41},1)", 4)] // When there is a sequence of target values, return last one
        [TestCase("MATCH(20,{25,38,40},1)", XLError.NoValueAvailable)] // Nothing found, even smallest value is greater than target
        [TestCase("MATCH(25,{20,TRUE,FALSE,38,40},1)", 1)] // If found value is <= target, return position of value, not subsequent types that are ignored
        [TestCase("MATCH(8, {FALSE;FALSE}, 1)", XLError.NoValueAvailable)] // Not even one value of target type
        [TestCase("MATCH(5, {1,2,3}, 1)", 3)] // If target value is greater than the last element of same type, return the position of the last element
        public void Match_from_ascending(string formula, object result)
        {
            var actual = XLWorkbook.EvaluateExpr(formula);
            Assert.AreEqual(result, actual);
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
            Assert.AreEqual(result, actual);
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
            Assert.AreEqual(result, actual);
        }

        [Test]
        public void Match_accepts_single_cell_as_values()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            sheet.Cell("A1").Value = 5;
            Assert.AreEqual(1, sheet.Evaluate("MATCH(5, A1)"));
        }

        [Test]
        public void Row()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Data");
            wb.AddWorksheet("Other");

            // If no argument, function uses the address of the cell that contains the formula
            Assert.AreEqual(60, ws.Cell("M60").SetFormulaA1("ROW()").Value);

            // With a reference, it returns the row number
            Assert.AreEqual(12, ws.Cell("A1").SetFormulaA1("ROW(C12)").Value);

            // If a full row reference to a single row is used, return the row number 
            Assert.AreEqual(40, ws.Cell("A2").SetFormulaA1("ROW(40:40)").Value);

            // Return a vertical array for multiple rows. Use SUM to verify content of an array since ROWS/COLUMNS don't work yet.
            Assert.AreEqual(4 + 5 + 6 + 7, ws.Cell("A3").SetFormulaA1("SUM(ROW(4:7))").Value);
            Assert.AreEqual(2 + 3 + 4, ws.Cell("A4").SetFormulaA1("SUM(ROW(C2:Z4))").Value);

            // Not contiguous range (multiple areas) returns #REF!
            Assert.AreEqual(XLError.CellReference, ws.Cell("A5").SetFormulaA1("ROW((D5:G10,I8:K12))").Value);

            // Invalid references return #REF!
            Assert.AreEqual(XLError.CellReference, ws.Cell("A6").SetFormulaA1("ROW(NonExistent!F10)").Value);

            // Return row number even for different worksheet
            Assert.AreEqual(14, ws.Cell("A7").SetFormulaA1("ROW(Other!E14)").Value);

            // Unexpected types return error
            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A8").SetFormulaA1("ROW(IF(TRUE,TRUE))").Value);
            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A9").SetFormulaA1("ROW(IF(TRUE,5))").Value);
            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A10").SetFormulaA1("ROW(IF(TRUE,\"G15\"))").Value);
            Assert.AreEqual(XLError.DivisionByZero, ws.Cell("A11").SetFormulaA1("ROW(#DIV/0!)").Value);

            // Properly works even in array formulas, where border between references and arrays blurs.
            ws.Range("A12:A13").FormulaArrayA1 = "ROW(2:3)";
            Assert.AreEqual(2, ws.Cell("A12").Value);
            Assert.AreEqual(3, ws.Cell("A13").Value);
        }

        [Test]
        public void Rows_Blank_ReturnsValueError()
        {
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("ROWS(IF(TRUE,,))"));
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
            Assert.AreEqual(1, XLWorkbook.EvaluateExpr($"ROWS({value})"));
        }

        [Test]
        public void Rows_Error_ReturnsError()
        {
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("ROWS(#DIV/0!)"));
        }

        [TestCase("{1}", 1)]
        [TestCase("{1;2;3}", 3)]
        [TestCase("{1,2,3,4;5,6,7,8;9,10,11,12}", 3)]
        [TestCase("{TRUE;#DIV/0!}", 2)]
        public void Rows_Arrays_ReturnsNumberOfRows(string array, int expectedColumnCount)
        {
            Assert.AreEqual(expectedColumnCount, XLWorkbook.EvaluateExpr($"ROWS({array})"));
        }

        [TestCase("C3", 1)]
        [TestCase("B3:E12", 10)]
        [TestCase("AA21:AC400", 380)]
        public void Rows_References_ReturnsNumberOfColumns(string range, int expectedColumnCount)
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            Assert.AreEqual(expectedColumnCount, sheet.Evaluate($"ROWS({range})"));
        }

        [Test]
        public void Rows_NonContiguousReferences_ReturnsReferenceError()
        {
            // Spec says #NULL!, but Excel says #REF!
            Assert.AreEqual(XLError.CellReference, XLWorkbook.EvaluateExpr("ROWS((A1,C3))"));
        }

        [Test]
        public void Vlookup()
        {
            // Range lookup false = exact match
            var value = ws.Evaluate("=VLOOKUP(3,Data!$B$2:$I$71,3,FALSE)");
            Assert.AreEqual("Central", value);

            value = ws.Evaluate("=VLOOKUP(DATE(2015,5,22),Data!C:I,7,FALSE)");
            Assert.AreEqual(63.68, value);

            value = ws.Evaluate(@"=VLOOKUP(""Central"",Data!D:E,2,FALSE)");
            Assert.AreEqual("Kivell", value);

            // Case insensitive lookup
            value = ws.Evaluate(@"=VLOOKUP(""central"",Data!D:E,2,FALSE)");
            Assert.AreEqual("Kivell", value);

            // Range lookup true = approximate match
            value = ws.Evaluate("=VLOOKUP(3,Data!$B$2:$I$71,8,TRUE)");
            Assert.AreEqual(179.64, value);

            value = ws.Evaluate("=VLOOKUP(3,Data!$B$2:$I$71,8)");
            Assert.AreEqual(179.64, value);

            value = ws.Evaluate("=VLOOKUP(3,Data!$B$2:$I$71,8,)");
            Assert.AreEqual(179.64, value);

            value = ws.Evaluate("=VLOOKUP(14.5,Data!$B$2:$I$71,8,TRUE)");
            Assert.AreEqual(174.65, value);

            value = ws.Evaluate("=VLOOKUP(50,Data!$B$2:$I$71,8,TRUE)");
            Assert.AreEqual(139.72, value);
        }

        [Test]
        public void Vlookup_ElementNotFound_ReturnsNotAvailableError()
        {
            // Value not present in the range for exact search
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"=VLOOKUP("""",Data!$B$2:$I$71,3,FALSE)"));
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"=VLOOKUP(50,Data!$B$2:$I$71,3,FALSE)"));

            // Value in approximate search that is lower than first element
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"=VLOOKUP(-1,Data!$B$2:$I$71,2,TRUE)"));
        }

        [Test]
        public void Vlookup_UnexpectedArguments()
        {
            // Lookup value can't be an error
            Assert.AreEqual(XLError.DivisionByZero, ws.Evaluate("=VLOOKUP(#DIV/0!,B2:I71,1)"));

            // Text value can't be over 255 chars
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate($"=VLOOKUP(\"{new string('A', 256)}\",B2:I71,1)"));

            // Range can only be array or a reference. If other type, it returns the error #N/A
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("=VLOOKUP(1,1,1)"));
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("=VLOOKUP(1,TRUE,1)"));

            // If range is a non-contiguous range, #N/A
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("=VLOOKUP(1,(B2:I5,B6:I10),1)"));

            // The column index must be at most the same as width of the range. It is 9 here, but range is 8 cell wide.
            Assert.AreEqual(XLError.CellReference, ws.Evaluate("=VLOOKUP(20,B2:I71,9,FALSE)"));
            // The column index must be at least 1. It is 0 here.
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("=VLOOKUP(20,B2:I71,0,FALSE)"));
        }

        [Test]
        public void Vlookup_ColumnIndexParameter_UsesValueSemantic()
        {
            // If column index is not a whole number, it is truncated, so here 1.9 is truncated to 1
            Assert.AreEqual(14.0, ws.Evaluate("=VLOOKUP(14,B2:I71,1.9)"));

            // Column index is evaluated using a VALUE semantic
            Assert.AreEqual(@"Jardine", ws.Evaluate("=VLOOKUP(3,B2:I71,\"2 5/2\")"));
        }

        [TestCase("\"TRUE\"")]
        [TestCase("1")]
        [TestCase("TRUE")]
        public void Vlookup_FlagParameter_CoercedToBoolean(string flagValue)
        {
            Assert.AreEqual(5.0, ws.Evaluate($"VLOOKUP(5,B2:I71,1,{flagValue})"));
        }

        [Test]
        public void Vlookup_BlankLookupValue_BehavesAsZero()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A1").InsertData(Enumerable.Range(-5, 10).Select(x => new object[] { x, $"Row with value {x}" }));

            var actual = worksheet.Evaluate("VLOOKUP(IF(TRUE,,),A1:B10,2)");

            Assert.AreEqual("Row with value 0", actual);
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
            Assert.AreEqual("Row 3", actual);
        }

        [Test]
        public void Vlookup_OnlyCellsWithDifferentType_ReturnsNotAvailable()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            Assert.AreEqual(XLError.NoValueAvailable, worksheet.Evaluate("VLOOKUP(1,A1,1,TRUE)"));
        }

        [Test]
        public void Vlookup_OnlyOneValueSurroundedByIgnoredTypes()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A3").Value = 5;

            Assert.AreEqual(5, worksheet.Evaluate("VLOOKUP(6,A1:A5,1,TRUE)"));
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

            Assert.AreEqual(3, worksheet.Evaluate("VLOOKUP(3,A1:A4,1,TRUE)"));
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
            Assert.AreEqual(7, actual);

            // If the last value is in the highest row, just return value outright
            actual = sheet.Evaluate("VLOOKUP(3, A2:B7, 2, TRUE)");
            Assert.AreEqual(7, actual);
        }

        [Test]
        public void Vlookup_CanSearchArrays()
        {
            Assert.AreEqual(2, XLWorkbook.EvaluateExpr("VLOOKUP(4, {1,2; 3,2; 5,3; 7,4}, 2)"));
        }
    }
}
