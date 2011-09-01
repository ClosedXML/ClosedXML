using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Excel
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class XLWorkbookTests
    {

        [TestMethod]
        public void WbNamedCell()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell");
            Assert.AreEqual("Test", wb.Cell("TestCell").GetString());
            Assert.AreEqual("Test", ws.Cell("TestCell").GetString());
        }

        [TestMethod]
        public void WbNamedCells()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell");
            ws.Cell(2, 1).SetValue("B").AddToNamed("Test2");
            var wbCells = wb.Cells("TestCell, Test2");
            Assert.AreEqual("Test", wbCells.First().GetString());
            Assert.AreEqual("B", wbCells.Last().GetString());

            var wsCells = ws.Cells("TestCell, Test2");
            Assert.AreEqual("Test", wsCells.First().GetString());
            Assert.AreEqual("B", wsCells.Last().GetString());

        }

        [TestMethod]
        public void WbNamedRange()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("A");
            ws.Cell(2, 1).SetValue("B");
            var original = ws.Range("A1:A2");
            original.AddToNamed("TestRange");
            Assert.AreEqual(original.RangeAddress.ToStringFixed(), wb.Range("TestRange").RangeAddress.ToString());
            Assert.AreEqual(original.RangeAddress.ToStringFixed(), ws.Range("TestRange").RangeAddress.ToString());
        }

        [TestMethod]
        public void WbNamedRanges()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("A");
            ws.Cell(2, 1).SetValue("B");
            ws.Cell(3, 1).SetValue("C").AddToNamed("Test2");
            var original = ws.Range("A1:A2");
            original.AddToNamed("TestRange");
            var wbRanges = wb.Ranges("TestRange, Test2");
            Assert.AreEqual(original.RangeAddress.ToStringFixed(), wbRanges.First().RangeAddress.ToString());
            Assert.AreEqual("$A$3:$A$3", wbRanges.Last().RangeAddress.ToStringFixed());

            var wsRanges = wb.Ranges("TestRange, Test2");
            Assert.AreEqual(original.RangeAddress.ToStringFixed(), wsRanges.First().RangeAddress.ToString());
            Assert.AreEqual("$A$3:$A$3", wsRanges.Last().RangeAddress.ToStringFixed());
        }

        [TestMethod]
        public void WbNamedRangesOneString()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            wb.NamedRanges.Add("TestRange", "Sheet1!$A$1,Sheet1!$A$3");

            var wbRanges = ws.Ranges("TestRange");
            Assert.AreEqual("$A$1:$A$1", wbRanges.First().RangeAddress.ToStringFixed());
            Assert.AreEqual("$A$3:$A$3", wbRanges.Last().RangeAddress.ToStringFixed());

            var wsRanges = ws.Ranges("TestRange");
            Assert.AreEqual("$A$1:$A$1", wsRanges.First().RangeAddress.ToStringFixed());
            Assert.AreEqual("$A$3:$A$3", wsRanges.Last().RangeAddress.ToStringFixed());
        }
    }
}
