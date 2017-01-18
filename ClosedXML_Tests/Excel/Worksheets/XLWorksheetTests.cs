﻿using System;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class XLWorksheetTests
    {
        [Test]
        public void ColumnCountTime()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            DateTime start = DateTime.Now;
            ws.ColumnCount();
            DateTime end = DateTime.Now;
            Assert.IsTrue((end - start).TotalMilliseconds < 500);
        }

        [Test]
        public void CopyConditionalFormatsCount()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().AddConditionalFormat().WhenContains("1").Fill.SetBackgroundColor(XLColor.Blue);
            IXLWorksheet ws2 = ws.CopyTo("Sheet2");
            Assert.AreEqual(1, ws2.ConditionalFormats.Count());
        }

        [Test]
        public void DeletingSheets1()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet3");
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets.Add("Sheet1", 1);

            wb.Worksheet("Sheet3").Delete();

            Assert.AreEqual("Sheet1", wb.Worksheet(1).Name);
            Assert.AreEqual("Sheet2", wb.Worksheet(2).Name);
            Assert.AreEqual(2, wb.Worksheets.Count);
        }

        [Test]
        public void InsertingSheets1()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet1");
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets.Add("Sheet3");

            Assert.AreEqual("Sheet1", wb.Worksheet(1).Name);
            Assert.AreEqual("Sheet2", wb.Worksheet(2).Name);
            Assert.AreEqual("Sheet3", wb.Worksheet(3).Name);
        }

        [Test]
        public void InsertingSheets2()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets.Add("Sheet1", 1);
            wb.Worksheets.Add("Sheet3");

            Assert.AreEqual("Sheet1", wb.Worksheet(1).Name);
            Assert.AreEqual("Sheet2", wb.Worksheet(2).Name);
            Assert.AreEqual("Sheet3", wb.Worksheet(3).Name);
        }

        [Test]
        public void InsertingSheets3()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet3");
            wb.Worksheets.Add("Sheet2", 1);
            wb.Worksheets.Add("Sheet1", 1);


            Assert.AreEqual("Sheet1", wb.Worksheet(1).Name);
            Assert.AreEqual("Sheet2", wb.Worksheet(2).Name);
            Assert.AreEqual("Sheet3", wb.Worksheet(3).Name);
        }

        [Test]
        public void MergedRanges()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("A1:B2").Merge();
            ws.Range("C1:D3").Merge();
            ws.Range("D2:E2").Merge();

            Assert.AreEqual(2, ws.MergedRanges.Count);
            Assert.AreEqual("A1:B2", ws.MergedRanges.First().RangeAddress.ToStringRelative());
            Assert.AreEqual("D2:E2", ws.MergedRanges.Last().RangeAddress.ToStringRelative());
        }

        [Test]
        public void RowCountTime()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            DateTime start = DateTime.Now;
            ws.RowCount();
            DateTime end = DateTime.Now;
            Assert.IsTrue((end - start).TotalMilliseconds < 500);
        }

        [Test]
        public void ShiftConditionalFormatsWhenInsertedRows()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell(5, 1).AddConditionalFormat().WhenIsTrue("=B5=0").Fill.SetBackgroundColor(XLColor.Blue);
            ws.Cell(5, 2).Value = 0;
            ws.Row(4).InsertRowsAbove(1);
            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            var format = ws.ConditionalFormats.First();
            Assert.AreEqual(6, format.Range.RangeAddress.FirstAddress.RowNumber);
            Assert.IsTrue(format.Values.Values.First().IsFormula);
            Assert.AreEqual("B6=0", format.Values.Values.First().Value);
        }

        [Test]
        public void ShiftConditionalFormatsWhenInsertedColumns()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell(1, 5).AddConditionalFormat().WhenIsTrue("=E2=0").Fill.SetBackgroundColor(XLColor.Blue);
            ws.Cell(2, 5).Value = 0;
            ws.Column(4).InsertColumnsBefore(1);
            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            var format = ws.ConditionalFormats.First();
            Assert.AreEqual(6, format.Range.RangeAddress.FirstAddress.ColumnNumber);
            Assert.IsTrue(format.Values.Values.First().IsFormula);
            Assert.AreEqual("F2=0", format.Values.Values.First().Value);
        }
    }
}