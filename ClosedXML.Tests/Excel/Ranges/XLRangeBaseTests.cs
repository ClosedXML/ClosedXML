using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Tests
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
            bool actual = range.IsEmpty(XLCellsUsedOptions.All);
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
            bool actual = range.IsEmpty(XLCellsUsedOptions.AllContents);
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
            bool actual = range.IsEmpty(XLCellsUsedOptions.All);
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

                Assert.AreEqual("D9:G11", ws.Range("B9:I11").Intersection(ws.Range("D4:G16")).ToString());
                Assert.AreEqual("E9:G11", ws.Range("E9:I11").Intersection(ws.Range("D4:G16")).ToString());
                Assert.AreEqual("E9:E9", ws.Cell("E9").AsRange().Intersection(ws.Range("D4:G16")).ToString());
                Assert.AreEqual("E9:E9", ws.Range("D4:G16").Intersection(ws.Cell("E9").AsRange()).ToString());

                XLRangeAddress rangeAddress;

                rangeAddress = (XLRangeAddress)ws.Cell("C3").AsRange().Intersection(ws.Cell("A1").AsRange());
                Assert.IsFalse(rangeAddress.IsValid);

                rangeAddress = (XLRangeAddress)ws.Cell("A1").AsRange().Intersection(ws.Cell("C3").AsRange());
                Assert.IsFalse(rangeAddress.IsValid);

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
            ws.Range("B2:E3").Clear(XLClearOptions.ConditionalFormats);

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            Assert.AreEqual("C4:D7", ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative());
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeAbove2()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("C3:D3").Clear(XLClearOptions.ConditionalFormats);

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            Assert.AreEqual("C4:D7", ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative());
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeBelow1()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("B7:E8").Clear(XLClearOptions.ConditionalFormats);

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            Assert.AreEqual("C3:D6", ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative());
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeBelow2()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("C7:D7").Clear(XLClearOptions.ConditionalFormats);

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            Assert.AreEqual("C3:D6", ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative());
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeRowInMiddle()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("C5:E5").Clear(XLClearOptions.ConditionalFormats);

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            Assert.AreEqual("C3:D4", ws.ConditionalFormats.First().Ranges.First().RangeAddress.ToStringRelative());
            Assert.AreEqual("C6:D7", ws.ConditionalFormats.First().Ranges.Last().RangeAddress.ToStringRelative());
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeColumnInMiddle()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:G4").AddConditionalFormat();
            ws.Range("E2:E4").Clear(XLClearOptions.ConditionalFormats);

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            Assert.AreEqual("C3:D4", ws.ConditionalFormats.First().Ranges.First().RangeAddress.ToStringRelative());
            Assert.AreEqual("F3:G4", ws.ConditionalFormats.First().Ranges.Last().RangeAddress.ToStringRelative());
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeContainsFormatWhole()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:G4").AddConditionalFormat();
            ws.Range("B2:G4").Clear(XLClearOptions.ConditionalFormats);

            Assert.AreEqual(0, ws.ConditionalFormats.Count());
        }

        [Test]
        public void NoClearConditionalFormattingsWhenRangePartiallySuperimposed()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:G4").AddConditionalFormat();
            ws.Range("C2:D3").Clear(XLClearOptions.ConditionalFormats);

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            Assert.AreEqual(1, ws.ConditionalFormats.Single().Ranges.Count);
            Assert.AreEqual("C3:G4", ws.ConditionalFormats.Single().Ranges.Single().RangeAddress.ToStringRelative());
        }

        [Test]
        public void RangesRemoveAllWithoutDispose()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var ranges = new XLRanges();
            ranges.Add(ws.Range("A1:A2"));
            ranges.Add(ws.Range("B1:B2"));
            var rangesCopy = ranges.ToList();

            ranges.RemoveAll(null, false);
            ws.FirstColumn().InsertColumnsBefore(1);

            Assert.AreEqual(0, ranges.Count);
            // if ranges were not disposed they addresses should change
            Assert.AreEqual("B1:B2", rangesCopy.First().RangeAddress.ToString());
            Assert.AreEqual("C1:C2", rangesCopy.Last().RangeAddress.ToString());
        }

        [Test]
        public void RangesRemoveAllByCriteria()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var ranges = new XLRanges();
            ranges.Add(ws.Range("A1:A2"));
            ranges.Add(ws.Range("B1:B3"));
            ranges.Add(ws.Range("C1:C4"));
            var otherRange = ws.Range("A3:D3");

            ranges.RemoveAll(r => r.Intersects(otherRange));

            Assert.AreEqual(1, ranges.Count);
            Assert.AreEqual("A1:A2", ranges.Single().RangeAddress.ToString());
        }

        [Test]
        public void XLRangesReturnsRangesInDeterministicOrder()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            var ws2 = wb.Worksheets.Add("Another sheet");

            var ranges = new XLRanges();
            ranges.Add(ws2.Range("F1:F12"));
            ranges.Add(ws1.Range("F12:F16"));
            ranges.Add(ws1.Range("B1:F2"));
            ranges.Add(ws2.Range("A13:B14"));
            ranges.Add(ws2.Range("E1:E2"));
            ranges.Add(ws1.Range("E1:H2"));
            ranges.Add(ws1.Range("G2:G13"));
            ranges.Add(ws1.Range("G20:G20"));

            var expectedRanges = new List<IXLRange>
            {
                ws1.Range("B1:F2"),
                ws1.Range("E1:H2"),
                ws1.Range("G2:G13"),
                ws1.Range("F12:F16"),
                ws1.Range("G20:G20"),

                ws2.Range("E1:E2"),
                ws2.Range("F1:F12"),
                ws2.Range("A13:B14"),
            };

            var actualRanges = ranges.ToList();

            Assert.AreEqual(expectedRanges.Count, actualRanges.Count);
            for (int i = 0; i < actualRanges.Count; i++)
            {
                Assert.AreEqual(expectedRanges[i], actualRanges[i]);
            }
        }

        [Test]
        public void ClearRangeRemovesSparklines()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.SparklineGroups.Add("B1:B3", "C1:E3");

            ws.Range("B1:C1").Clear(XLClearOptions.All);
            ws.Range("B2:C2").Clear(XLClearOptions.Sparklines);

            Assert.AreEqual(1, ws.SparklineGroups.Single().Count());
            Assert.IsFalse(ws.Cell("B1").HasSparkline);
            Assert.IsFalse(ws.Cell("B2").HasSparkline);
            Assert.IsTrue(ws.Cell("B3").HasSparkline);
        }

        [TestCase("B2:G7", "D4:E5", true, "B2:G3,B4:C5,D4:E5,F4:G5,B6:G7")]
        [TestCase("B2:G7", "D4:E5", false, "B2:G3,B4:C5,F4:G5,B6:G7")]
        [TestCase("B2:G7", "B2:G7", true, "B2:G7")]
        [TestCase("B2:G7", "B2:G7", false, "")]
        [TestCase("B2:G7", "A1:H8", true, "B2:G7")]
        [TestCase("B2:G7", "A1:H8", false, "")]
        [TestCase("B2:G7", "A1:B2", true, "B2:B2,C2:G2,B3:G7")]
        [TestCase("B2:G7", "A1:B2", false, "C2:G2,B3:G7")]
        [TestCase("B2:G7", "E4:J5", true, "B2:G3,B4:D5,E4:G5,B6:G7")]
        [TestCase("B2:G7", "E4:J5", false, "B2:G3,B4:D5,B6:G7")]
        [TestCase("B2:G7", "A11:H18", true, "B2:G7")]
        [TestCase("B2:G7", "A11:H18", false, "B2:G7")]
        [TestCase("B2:G7", "A1:H1", true, "B2:G7")]
        [TestCase("B2:G7", "A1:A12", true, "B2:G7")]
        [TestCase("B2:G7", "A8:H8", true, "B2:G7")]
        [TestCase("B2:G7", "H1:H8", true, "B2:G7")]
        public void CanSplitRange(string rangeAddress, string splitBy, bool includeIntersection, string expectedResult)
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var range = ws.Range(rangeAddress) as XLRange;
            var splitter = ws.Range(splitBy);

            var result = range.Split(splitter.RangeAddress, includeIntersection);

            var actualAddresses = string.Join(",", result.Select(r => r.RangeAddress.ToString()));

            Assert.AreEqual(expectedResult, actualAddresses);
        }
    }
}
