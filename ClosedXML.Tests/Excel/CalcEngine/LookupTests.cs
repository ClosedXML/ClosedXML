// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using ClosedXML.Excel.CalcEngine.Exceptions;
using NUnit.Framework;
using System;

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
        public void Hlookup()
        {
            // Range lookup false
            var value = ws.Evaluate(@"=HLOOKUP(""Total"",Data!$B$2:$I$71,4,FALSE)");
            Assert.AreEqual(179.64, value);
        }

        [Test]
        public void Hyperlink()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            var cell = ws.Cell("B3");
            cell.FormulaA1 = "HYPERLINK(\"http://github.com/ClosedXML/ClosedXML\")";
            Assert.AreEqual("http://github.com/ClosedXML/ClosedXML", cell.Value);
            Assert.True(cell.HasHyperlink);
            Assert.AreEqual("http://github.com/ClosedXML/ClosedXML", cell.GetHyperlink().ExternalAddress.ToString());

            cell = ws.Cell("B4");
            cell.FormulaA1 = "HYPERLINK(\"mailto:jsmith@github.com\", \"jsmith@github.com\")";
            Assert.AreEqual("jsmith@github.com", cell.Value);
            Assert.True(cell.HasHyperlink);
            Assert.AreEqual("mailto:jsmith@github.com", cell.GetHyperlink().ExternalAddress.ToString());
        }

        [Test]
        public void Index()
        {
            Assert.AreEqual("Kivell", ws.Evaluate(@"=INDEX(B2:J12, 3, 4)"));

            // We don't support optional parameter fully here yet.
            // Supposedly, if you omit e.g. the row number, then ROW() of the calling cell should be assumed
            // Assert.AreEqual("Gill", ws.Evaluate(@"=INDEX(B2:J12, , 4)"));

            Assert.AreEqual("Rep", ws.Evaluate(@"=INDEX(B2:I2, 4)"));

            Assert.AreEqual(3, ws.Evaluate(@"=INDEX(B2:B20, 4)"));
            Assert.AreEqual(3, ws.Evaluate(@"=INDEX(B2:B20, 4, 1)"));
            Assert.AreEqual(3, ws.Evaluate(@"=INDEX(B2:B20, 4, )"));

            Assert.AreEqual("Rep", ws.Evaluate(@"=INDEX(B2:J2, 1, 4)"));
            Assert.AreEqual("Rep", ws.Evaluate(@"=INDEX(B2:J2, , 4)"));
        }

        [Test]
        public void Index_Exceptions()
        {
            Assert.Throws<CellReferenceException>(() => ws.Evaluate(@"INDEX(B2:I10, 20, 1)"));
            Assert.Throws<CellReferenceException>(() => ws.Evaluate(@"INDEX(B2:I10, 1, 10)"));
            Assert.Throws<CellReferenceException>(() => ws.Evaluate(@"INDEX(B2:I2, 10)"));
            Assert.Throws<CellReferenceException>(() => ws.Evaluate(@"INDEX(B2:I2, 4, 1)"));
            Assert.Throws<CellReferenceException>(() => ws.Evaluate(@"INDEX(B2:I2, 4, )"));
            Assert.Throws<CellReferenceException>(() => ws.Evaluate(@"INDEX(B2:B10, 20)"));
            Assert.Throws<CellReferenceException>(() => ws.Evaluate(@"INDEX(B2:B10, 20, )"));
            Assert.Throws<CellReferenceException>(() => ws.Evaluate(@"INDEX(B2:B10, , 4)"));
        }

        [Test]
        public void Match()
        {
            Object value;
            value = ws.Evaluate(@"=MATCH(""Rep"", B2:I2, 0)");
            Assert.AreEqual(4, value);

            value = ws.Evaluate(@"=MATCH(""Rep"", A2:Z2, 0)");
            Assert.AreEqual(5, value);

            value = ws.Evaluate(@"=MATCH(""REP"", B2:I2, 0)");
            Assert.AreEqual(4, value);

            value = ws.Evaluate(@"=MATCH(95, B3:I3, 0)");
            Assert.AreEqual(6, value);

            value = ws.Evaluate(@"=MATCH(DATE(2015,1,6), B3:I3, 0)");
            Assert.AreEqual(2, value);

            value = ws.Evaluate(@"=MATCH(1.99, 3:3, 0)");
            Assert.AreEqual(8, value);

            value = ws.Evaluate(@"=MATCH(43, B:B, 0)");
            Assert.AreEqual(45, value);

            value = ws.Evaluate(@"=MATCH(""cENtraL"", D3:D45, 0)");
            Assert.AreEqual(2, value);

            value = ws.Evaluate(@"=MATCH(4.99, H:H, 0)");
            Assert.AreEqual(5, value);

            value = ws.Evaluate(@"=MATCH(""Rapture"", B2:I2, 1)");
            Assert.AreEqual(2, value);

            value = ws.Evaluate(@"=MATCH(22.5, B3:B45, 1)");
            Assert.AreEqual(22, value);

            value = ws.Evaluate(@"=MATCH(""Rep"", B2:I2)");
            Assert.AreEqual(4, value);

            value = ws.Evaluate(@"=MATCH(""Rep"", B2:I2, 1)");
            Assert.AreEqual(4, value);

            value = ws.Evaluate(@"=MATCH(40, G3:G6, -1)");
            Assert.AreEqual(2, value);
        }

        [Test]
        public void Match_Exceptions()
        {
            Assert.Throws<CellValueException>(() => ws.Evaluate(@"=MATCH(""Rep"", B2:I5)"));
            Assert.Throws<NoValueAvailableException>(() => ws.Evaluate(@"=MATCH(""Dummy"", B2:I2, 0)"));
            Assert.Throws<NoValueAvailableException>(() => ws.Evaluate(@"=MATCH(4.5,B3:B45,-1)"));
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
        }

        [Test]
        public void Vlookup()
        {
            // Range lookup false
            var value = ws.Evaluate("=VLOOKUP(3,Data!$B$2:$I$71,3,FALSE)");
            Assert.AreEqual("Central", value);

            value = ws.Evaluate("=VLOOKUP(DATE(2015,5,22),Data!C:I,7,FALSE)");
            Assert.AreEqual(63.68, value);

            value = ws.Evaluate(@"=VLOOKUP(""Central"",Data!D:E,2,FALSE)");
            Assert.AreEqual("Kivell", value);

            // Case insensitive lookup
            value = ws.Evaluate(@"=VLOOKUP(""central"",Data!D:E,2,FALSE)");
            Assert.AreEqual("Kivell", value);

            // Range lookup true
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
        public void Vlookup_Exceptions()
        {
            Assert.Throws<NoValueAvailableException>(() => ws.Evaluate(@"=VLOOKUP("""",Data!$B$2:$I$71,3,FALSE)"));
            Assert.Throws<NoValueAvailableException>(() => ws.Evaluate(@"=VLOOKUP(50,Data!$B$2:$I$71,3,FALSE)"));
            Assert.Throws<NoValueAvailableException>(() => ws.Evaluate(@"=VLOOKUP(-1,Data!$B$2:$I$71,2,TRUE)"));

            Assert.Throws<CellReferenceException>(() => ws.Evaluate(@"=VLOOKUP(20,Data!$B$2:$I$71,9,FALSE)"));
        }
    }
}
