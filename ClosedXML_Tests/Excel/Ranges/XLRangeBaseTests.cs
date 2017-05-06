using System;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class XLRangeBaseTests
    {
        [Test]
        public void IsEmpty1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            IXLRange range = ws.Range("A1:B2");
            bool actual = range.IsEmpty();
            bool expected = true;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void IsEmpty2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            IXLRange range = ws.Range("A1:B2");
            bool actual = range.IsEmpty(true);
            bool expected = true;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void IsEmpty3()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            IXLRange range = ws.Range("A1:B2");
            bool actual = range.IsEmpty();
            bool expected = true;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void IsEmpty4()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            IXLRange range = ws.Range("A1:B2");
            bool actual = range.IsEmpty(false);
            bool expected = true;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void IsEmpty5()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            IXLRange range = ws.Range("A1:B2");
            bool actual = range.IsEmpty(true);
            bool expected = false;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void IsEmpty6()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Value = "X";
            IXLRange range = ws.Range("A1:B2");
            bool actual = range.IsEmpty();
            bool expected = false;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void SingleCell()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).Value = "Hello World!";
            wb.NamedRanges.Add("SingleCell", "Sheet1!$A$1");
            IXLRange range = wb.Range("SingleCell");
            Assert.AreEqual(1, range.CellsUsed().Count());
            Assert.AreEqual("Hello World!", range.CellsUsed().Single().GetString());
        }

        [Test]
        public void TableRange()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            IXLRangeColumn rangeColumn = ws.Column(1).Column(1, 4);
            rangeColumn.Cell(1).Value = "FName";
            rangeColumn.Cell(2).Value = "John";
            rangeColumn.Cell(3).Value = "Hank";
            rangeColumn.Cell(4).Value = "Dagny";
            IXLTable table = rangeColumn.CreateTable();
            wb.NamedRanges.Add("FNameColumn", String.Format("{0}[{1}]", table.Name, "FName"));

            IXLRange namedRange = wb.Range("FNameColumn");
            Assert.AreEqual(3, namedRange.Cells().Count());
            Assert.IsTrue(
                namedRange.CellsUsed().Select(cell => cell.GetString()).SequenceEqual(new[] {"John", "Hank", "Dagny"}));
        }

        [Test]
        public void WsNamedCell()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell", XLScope.Worksheet);
            Assert.AreEqual("Test", ws.Cell("TestCell").GetString());
        }

        [Test]
        public void WsNamedCells()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell", XLScope.Worksheet);
            ws.Cell(2, 1).SetValue("B");
            IXLCells cells = ws.Cells("TestCell, A2");
            Assert.AreEqual("Test", cells.First().GetString());
            Assert.AreEqual("B", cells.Last().GetString());
        }

        [Test]
        public void WsNamedRange()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("A");
            ws.Cell(2, 1).SetValue("B");
            IXLRange original = ws.Range("A1:A2");
            original.AddToNamed("TestRange", XLScope.Worksheet);
            IXLRange named = ws.Range("TestRange");
            Assert.AreEqual(original.RangeAddress.ToStringFixed(), named.RangeAddress.ToString());
        }

        [Test]
        public void WsNamedRanges()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("A");
            ws.Cell(2, 1).SetValue("B");
            ws.Cell(3, 1).SetValue("C");
            IXLRange original = ws.Range("A1:A2");
            original.AddToNamed("TestRange", XLScope.Worksheet);
            IXLRanges namedRanges = ws.Ranges("TestRange, A3");
            Assert.AreEqual(original.RangeAddress.ToStringFixed(), namedRanges.First().RangeAddress.ToString());
            Assert.AreEqual("$A$3:$A$3", namedRanges.Last().RangeAddress.ToStringFixed());
        }

        [Test]
        public void WsNamedRangesOneString()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.NamedRanges.Add("TestRange", "Sheet1!$A$1,Sheet1!$A$3");
            IXLRanges namedRanges = ws.Ranges("TestRange");

            Assert.AreEqual("$A$1:$A$1", namedRanges.First().RangeAddress.ToStringFixed());
            Assert.AreEqual("$A$3:$A$3", namedRanges.Last().RangeAddress.ToStringFixed());
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeAbove1()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("B2:E3").Clear(XLClearOptions.Formats);

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            Assert.AreEqual("C4:D7", ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative());
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeAbove2()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("C3:D3").Clear(XLClearOptions.Formats);

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            Assert.AreEqual("C4:D7", ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative());
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeBelow1()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("B7:E8").Clear(XLClearOptions.Formats);

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            Assert.AreEqual("C3:D6", ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative());
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeBelow2()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("C7:D7").Clear(XLClearOptions.Formats);

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            Assert.AreEqual("C3:D6", ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative());
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeRowInMiddle()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("C5:E5").Clear(XLClearOptions.Formats);

            Assert.AreEqual(2, ws.ConditionalFormats.Count());
            Assert.IsTrue(ws.ConditionalFormats.Any(x=>x.Range.RangeAddress.ToStringRelative() == "C3:D4"));
            Assert.IsTrue(ws.ConditionalFormats.Any(x=>x.Range.RangeAddress.ToStringRelative() == "C6:D7"));
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeColumnInMiddle()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:G4").AddConditionalFormat();
            ws.Range("E2:E4").Clear(XLClearOptions.Formats);

            Assert.AreEqual(2, ws.ConditionalFormats.Count());
            Assert.IsTrue(ws.ConditionalFormats.Any(x=>x.Range.RangeAddress.ToStringRelative() == "C3:D4"));
            Assert.IsTrue(ws.ConditionalFormats.Any(x=>x.Range.RangeAddress.ToStringRelative() == "F3:G4"));
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeContainsFormatWhole()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:G4").AddConditionalFormat();
            ws.Range("B2:G4").Clear(XLClearOptions.Formats);

            Assert.AreEqual(0, ws.ConditionalFormats.Count());
        }

        [Test]
        public void NoClearConditionalFormattingsWhenRangePartiallySuperimposed()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:G4").AddConditionalFormat();
            ws.Range("C2:D3").Clear(XLClearOptions.Formats);

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            Assert.AreEqual("C3:G4", ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative());
        }

        //[Test]
        //public void WsNamedRangeLiteral()
        //{
        //    var wb = new XLWorkbook();
        //    var ws = wb.Worksheets.Add("Sheet1");
        //    ws.NamedRanges.Add("TestRange", "\"Hello\"");
        //    using (MemoryStream memoryStream = new MemoryStream())
        //    {
        //        wb.SaveAs(memoryStream, true);
        //        var wb2 = new XLWorkbook(memoryStream);
        //        var text = wb2.Worksheet("Sheet1").NamedRanges.First()
        //        memoryStream.Close();
        //    }


        //}
    }
}
