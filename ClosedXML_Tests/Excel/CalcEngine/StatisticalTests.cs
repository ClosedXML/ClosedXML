using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML_Tests.Excel.CalcEngine
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

            Assert.That(() => ws.Evaluate("AVERAGE(D3:D45)"), Throws.Exception);
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

            value = workbook.Evaluate(@"=COUNTBLANK(E3:E45)").CastTo<int>();
            Assert.AreEqual(0, value);
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

        [OneTimeTearDown]
        public void Dispose()
        {
            workbook.Dispose();
        }

        [OneTimeSetUp]
        public void Init()
        {
            // Make sure tests run on a deterministic culture
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            workbook = SetupWorkbook();
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
            Assert.That(() => ws.Evaluate(@"=STDEV(D3:D45)"), Throws.Exception);

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
            Assert.That(() => ws.Evaluate(@"=STDEVP(D3:D45)"), Throws.Exception);

            value = ws.Evaluate(@"=STDEVP(H3:H45)").CastTo<double>();
            Assert.AreEqual(46.79135458, value, tolerance);

            value = ws.Evaluate(@"=STDEVP(H:H)").CastTo<double>();
            Assert.AreEqual(46.79135458, value, tolerance);

            value = workbook.Evaluate(@"=STDEVP(Data!H:H)").CastTo<double>();
            Assert.AreEqual(46.79135458, value, tolerance);
        }

        [Test]
        public void Var()
        {
            var ws = workbook.Worksheets.First();
            double value;
            Assert.That(() => ws.Evaluate(@"=VAR(D3:D45)"), Throws.Exception);

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
            Assert.That(() => ws.Evaluate(@"=VARP(D3:D45)"), Throws.Exception);

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
                .InsertTable(data, "Table1");

            return wb;
        }
    }
}
