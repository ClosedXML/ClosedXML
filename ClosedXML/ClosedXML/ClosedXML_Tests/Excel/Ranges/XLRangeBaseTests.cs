using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System;
using System.IO;

namespace ClosedXML_Tests
{
    [TestClass()]
    public class XLRangeBaseTests
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
                
                var namedRange = wb.Range( "FNameColumn" );
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
            var range = wb.Range( "SingleCell" );
            Assert.AreEqual( 1, range.CellsUsed().Count() );
            Assert.AreEqual("Hello World!", range.CellsUsed().Single().GetString());
        }

        [TestMethod]
        public void WsNamedCell()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell", XLScope.Worksheet);
            Assert.AreEqual("Test", ws.Cell("TestCell").GetString());
        }

        [TestMethod]
        public void WsNamedCells()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell", XLScope.Worksheet);
            ws.Cell(2, 1).SetValue("B");
            var cells = ws.Cells("TestCell, A2");
            Assert.AreEqual("Test", cells.First().GetString());
            Assert.AreEqual("B", cells.Last().GetString());
        }

        [TestMethod]
        public void WsNamedRange()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("A");
            ws.Cell(2, 1).SetValue("B");
            var original = ws.Range("A1:A2");
            original.AddToNamed("TestRange", XLScope.Worksheet);
            var named = ws.Range("TestRange");
            Assert.AreEqual(original.RangeAddress.ToStringFixed(), named.RangeAddress.ToString());
        }

        [TestMethod]
        public void WsNamedRanges()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("A");
            ws.Cell(2, 1).SetValue("B");
            ws.Cell(3, 1).SetValue("C");
            var original = ws.Range("A1:A2");
            original.AddToNamed("TestRange", XLScope.Worksheet);
            var namedRanges = ws.Ranges("TestRange, A3");
            Assert.AreEqual(original.RangeAddress.ToStringFixed(), namedRanges.First().RangeAddress.ToString());
            Assert.AreEqual("$A$3:$A$3", namedRanges.Last().RangeAddress.ToStringFixed());
        }

        [TestMethod]
        public void WsNamedRangesOneString()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.NamedRanges.Add("TestRange", "Sheet1!$A$1,Sheet1!$A$3");
            var namedRanges = ws.Ranges("TestRange");

            Assert.AreEqual("$A$1:$A$1", namedRanges.First().RangeAddress.ToStringFixed());
            Assert.AreEqual("$A$3:$A$3", namedRanges.Last().RangeAddress.ToStringFixed());
        }

        //[TestMethod]
        //public void WsNamedRangeLiteral()
        //{
        //    var wb = new XLWorkbook();
        //    var ws = wb.Worksheets.Add("Sheet1");
        //    ws.NamedRanges.Add("TestRange", "\"Hello\"");
        //    using (MemoryStream memoryStream = new MemoryStream())
        //    {
        //        wb.SaveAs(memoryStream);
        //        var wb2 = new XLWorkbook(memoryStream);
        //        var text = wb2.Worksheet("Sheet1").NamedRanges.First()
        //        memoryStream.Close();
        //    }
            
            
        //}
    }
}
