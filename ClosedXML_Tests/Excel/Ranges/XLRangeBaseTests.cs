using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

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
                namedRange.CellsUsed().Select(cell => cell.GetString()).SequenceEqual(new[] { "John", "Hank", "Dagny" }));
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

        [Test]
        public void GrowRange()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                Assert.AreEqual("A1:B2", ws.Cell("A1").AsRange().Grow().RangeAddress.ToString());
                Assert.AreEqual("A1:B3", ws.Cell("A2").AsRange().Grow().RangeAddress.ToString());
                Assert.AreEqual("A1:C2", ws.Cell("B1").AsRange().Grow().RangeAddress.ToString());

                Assert.AreEqual("E4:G6", ws.Cell("F5").AsRange().Grow().RangeAddress.ToString());
                Assert.AreEqual("D3:H7", ws.Cell("F5").AsRange().Grow(2).RangeAddress.ToString());
                Assert.AreEqual("A1:DB105", ws.Cell("F5").AsRange().Grow(100).RangeAddress.ToString());
            }
        }

        [Test]
        public void ShrinkRange()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                Assert.Null(ws.Cell("A1").AsRange().Shrink());
                Assert.Null(ws.Range("B2:C3").Shrink());
                Assert.AreEqual("C3:C3", ws.Range("B2:D4").Shrink().RangeAddress.ToString());
                Assert.AreEqual("K11:P16", ws.Range("A1:Z26").Shrink(10).RangeAddress.ToString());

                // Grow and shrink back
                Assert.AreEqual("Z26:Z26", ws.Cell("Z26").AsRange().Grow(10).Shrink(10).RangeAddress.ToString());
            }
        }

        [Test]
        public void Intersection()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                Assert.AreEqual("D9:G11", ws.Range("B9:I11").Intersection(ws.Range("D4:G16")).RangeAddress.ToString());
                Assert.AreEqual("E9:G11", ws.Range("E9:I11").Intersection(ws.Range("D4:G16")).RangeAddress.ToString());
                Assert.AreEqual("E9:E9", ws.Cell("E9").AsRange().Intersection(ws.Range("D4:G16")).RangeAddress.ToString());
                Assert.AreEqual("E9:E9", ws.Range("D4:G16").Intersection(ws.Cell("E9").AsRange()).RangeAddress.ToString());

                Assert.Null(ws.Cell("A1").AsRange().Intersection(ws.Cell("C3").AsRange()));

                Assert.Null(ws.Range("A1:C3").Intersection(null));

                var otherWs = wb.AddWorksheet("Sheet2");
                Assert.Null(ws.Intersection(otherWs));
                Assert.Null(ws.Cell("A1").AsRange().Intersection(otherWs.Cell("A2").AsRange()));
            }
        }

        [Test]
        public void Union()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                Assert.AreEqual(64, ws.Range("B9:I11").Union(ws.Range("D4:G16")).Count());
                Assert.AreEqual(58, ws.Range("E9:I11").Union(ws.Range("D4:G16")).Count());
                Assert.AreEqual(52, ws.Cell("E9").AsRange().Union(ws.Range("D4:G16")).Count());
                Assert.AreEqual(52, ws.Range("D4:G16").Union(ws.Cell("E9").AsRange()).Count());

                Assert.AreEqual(2, ws.Cell("A1").AsRange().Union(ws.Cell("C3").AsRange()).Count());

                Assert.AreEqual(9, ws.Range("A1:C3").Union(null).Count());

                var otherWs = wb.AddWorksheet("Sheet2");
                Assert.False(ws.Union(otherWs).Any());
                Assert.False(ws.Cell("A1").AsRange().Union(otherWs.Cell("A2").AsRange()).Any());
            }
        }

        [Test]
        public void Difference()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                Assert.AreEqual(12, ws.Range("B9:I11").Difference(ws.Range("D4:G16")).Count());
                Assert.AreEqual(6, ws.Range("E9:I11").Difference(ws.Range("D4:G16")).Count());
                Assert.AreEqual(0, ws.Cell("E9").AsRange().Difference(ws.Range("D4:G16")).Count());
                Assert.AreEqual(51, ws.Range("D4:G16").Difference(ws.Cell("E9").AsRange()).Count());

                Assert.AreEqual(1, ws.Cell("A1").AsRange().Difference(ws.Cell("C3").AsRange()).Count());

                Assert.AreEqual(9, ws.Range("A1:C3").Difference(null).Count());

                var otherWs = wb.AddWorksheet("Sheet2");
                Assert.False(ws.Difference(otherWs).Any());
                Assert.False(ws.Cell("A1").AsRange().Difference(otherWs.Cell("A2").AsRange()).Any());
            }
        }

        [Test]
        public void SurroundingCells()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                Assert.AreEqual(3, ws.FirstCell().AsRange().SurroundingCells().Count());
                Assert.AreEqual(8, ws.Cell("C3").AsRange().SurroundingCells().Count());
                Assert.AreEqual(16, ws.Range("C3:D6").AsRange().SurroundingCells().Count());

                Assert.AreEqual(0, ws.Range("C3:D6").AsRange().SurroundingCells(c => !c.IsEmpty()).Count());
            }
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
            Assert.IsTrue(ws.ConditionalFormats.Any(x => x.Range.RangeAddress.ToStringRelative() == "C3:D4"));
            Assert.IsTrue(ws.ConditionalFormats.Any(x => x.Range.RangeAddress.ToStringRelative() == "C6:D7"));
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeColumnInMiddle()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:G4").AddConditionalFormat();
            ws.Range("E2:E4").Clear(XLClearOptions.Formats);

            Assert.AreEqual(2, ws.ConditionalFormats.Count());
            Assert.IsTrue(ws.ConditionalFormats.Any(x => x.Range.RangeAddress.ToStringRelative() == "C3:D4"));
            Assert.IsTrue(ws.ConditionalFormats.Any(x => x.Range.RangeAddress.ToStringRelative() == "F3:G4"));
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
    }
}
