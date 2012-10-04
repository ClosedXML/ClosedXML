using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System;

namespace ClosedXML_Tests
{
    [TestClass]
    public class XLCellTest
    {
        [TestMethod]
        public void IsEmpty1()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            var actual = cell.IsEmpty();
            var expected = true;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void IsEmpty2()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            var actual = cell.IsEmpty(true);
            var expected = true;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void IsEmpty3()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            var actual = cell.IsEmpty();
            var expected = true;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void IsEmpty4()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            var actual = cell.IsEmpty(false);
            var expected = true;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void IsEmpty5()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            var actual = cell.IsEmpty(true);
            var expected = false;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void IsEmpty6()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Value = "X";
            var actual = cell.IsEmpty();
            var expected = false;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ValueSetToEmptyString()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Value = new DateTime(2000, 1, 2);
            cell.Value = String.Empty;
            var actual = cell.GetString();
            var expected = String.Empty;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ValueSetToNull()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Value = new DateTime(2000, 1, 2);
            cell.Value = null;
            var actual = cell.GetString();
            var expected = String.Empty;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void InsertData1()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var range = ws.Cell(2, 2).InsertData(new[] {"a", "b", "c"});
            Assert.AreEqual("'Sheet1'!B2:B4", range.ToString());
        }

        [TestMethod]
        public void CellsUsed()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Cell(1, 1);
            ws.Cell(2, 2);
            var count = ws.Range("A1:B2").CellsUsed().Count();
            Assert.AreEqual(0, count);
        }

        [TestMethod]
        public void TryGetValue_TimeSpan_Good()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            TimeSpan outValue;
            var timeSpan = new TimeSpan(1, 1, 1);
            var success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(timeSpan, outValue);
        }

        [TestMethod]
        public void TryGetValue_TimeSpan_BadString()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            TimeSpan outValue;
            var timeSpan = "ABC";
            var success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out outValue);
            Assert.IsFalse(success);
        }

        [TestMethod]
        public void TryGetValue_TimeSpan_GoodString()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            TimeSpan outValue;
            var timeSpan = new TimeSpan(1, 1, 1);
            var success = ws.Cell("A1").SetValue(timeSpan.ToString()).TryGetValue(out outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(timeSpan, outValue);
        }
    }
}