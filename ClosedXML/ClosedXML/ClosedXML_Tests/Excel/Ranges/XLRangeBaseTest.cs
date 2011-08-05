using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System;

namespace ClosedXML_Tests
{
    [TestClass()]
    public class XLRangeBaseTest
    {
        [TestMethod()]
        public void IsEmpty1()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            var range = ws.Range("A1:B2");
            var actual = range.IsEmpty();
            var expected = true;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void IsEmpty2()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            var range = ws.Range("A1:B2");
            var actual = range.IsEmpty(true);
            var expected = true;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void IsEmpty3()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            var range = ws.Range("A1:B2");
            var actual = range.IsEmpty();
            var expected = true;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void IsEmpty4()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            var range = ws.Range("A1:B2");
            var actual = range.IsEmpty(false);
            var expected = true;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void IsEmpty5()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            var range = ws.Range("A1:B2");
            var actual = range.IsEmpty(true);
            var expected = false;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void IsEmpty6()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Value = "X";
            var range = ws.Range("A1:B2");
            var actual = range.IsEmpty();
            var expected = false;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TableRange()
        {
                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Sheet1");
                var rangeColumn = ws.Column(1).Column(1, 4);
                rangeColumn.Cell(1).Value = "FName";
                rangeColumn.Cell(2).Value = "John";
                rangeColumn.Cell(3).Value = "Hank";
                rangeColumn.Cell(4).Value = "Dagny";
                var table = rangeColumn.CreateTable();
                wb.NamedRanges.Add("FNameColumn", String.Format("{0}[{1}]", table.Name, "FName"));
                
                var namedRange = wb.NamedRange( "FNameColumn" ).Range;
                Assert.AreEqual(3, namedRange.Cells().Count());
                Assert.IsTrue(namedRange.CellsUsed().Select(cell => cell.GetString()).SequenceEqual(new[] { "John", "Hank", "Dagny" }));
        }

        [TestMethod]
        public void SingleCell()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).Value = "Hello World!";
            wb.NamedRanges.Add("SingleCell", "Sheet1!$A$1");
            var range = wb.NamedRange( "SingleCell" ).Range;
            Assert.AreEqual( 1, range.CellsUsed().Count() );
            Assert.AreEqual("Hello World!", range.CellsUsed().Single().GetString());
        }

    }
}
