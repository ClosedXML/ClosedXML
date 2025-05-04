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
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            IXLRange range = ws.Range("A1:B2");
            bool actual = range.IsEmpty(XLCellsUsedOptions.All);
            bool expected = true;
            Assert.That(actual, Is.EqualTo(expected));
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
            Assert.That(actual, Is.EqualTo(expected));
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
            Assert.That(actual, Is.EqualTo(expected));
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
            Assert.That(actual, Is.EqualTo(expected));
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
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void SingleCell()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).Value = "Hello World!";
            wb.DefinedNames.Add("SingleCell", "Sheet1!$A$1");
            IXLRange range = wb.Range("SingleCell");
            Assert.Multiple(() =>
            {
                Assert.That(range.CellsUsed().Count(), Is.EqualTo(1));
                Assert.That(range.CellsUsed().Single().GetText(), Is.EqualTo("Hello World!"));
            });
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
            wb.DefinedNames.Add("FNameColumn", String.Format("{0}[{1}]", table.Name, "FName"));

            IXLRange namedRange = wb.Range("FNameColumn");
            Assert.Multiple(() =>
            {
                Assert.That(namedRange.Cells().Count(), Is.EqualTo(3));
                Assert.That(
                    namedRange.CellsUsed().Select(cell => cell.GetText()).SequenceEqual(new[] { "John", "Hank", "Dagny" }),
                    Is.True);
            });
        }

        [Test]
        public void WsNamedCell()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell", XLScope.Worksheet);
            Assert.That(ws.Cell("TestCell").GetText(), Is.EqualTo("Test"));
        }

        [Test]
        public void WsNamedCells()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell", XLScope.Worksheet);
            ws.Cell(2, 1).SetValue("B");
            IXLCells cells = ws.Cells("TestCell, A2");
            Assert.Multiple(() =>
            {
                Assert.That(cells.First().GetText(), Is.EqualTo("Test"));
                Assert.That(cells.Last().GetText(), Is.EqualTo("B"));
            });
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
            Assert.That(named.RangeAddress.ToString(), Is.EqualTo(original.RangeAddress.ToStringFixed()));
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
            Assert.Multiple(() =>
            {
                Assert.That(namedRanges.First().RangeAddress.ToString(), Is.EqualTo(original.RangeAddress.ToStringFixed()));
                Assert.That(namedRanges.Last().RangeAddress.ToStringFixed(), Is.EqualTo("$A$3:$A$3"));
            });
        }

        [Test]
        public void WsNamedRangesOneString()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.DefinedNames.Add("TestRange", "Sheet1!$A$1,Sheet1!$A$3");
            IXLRanges namedRanges = ws.Ranges("TestRange");

            Assert.That(namedRanges.First().RangeAddress.ToStringFixed(), Is.EqualTo("$A$1:$A$1"));
            Assert.That(namedRanges.Last().RangeAddress.ToStringFixed(), Is.EqualTo("$A$3:$A$3"));
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
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            Assert.That(ws.Cell("A1").AsRange().Grow().RangeAddress.ToString(), Is.EqualTo("A1:B2"));
            Assert.That(ws.Cell("A2").AsRange().Grow().RangeAddress.ToString(), Is.EqualTo("A1:B3"));
            Assert.That(ws.Cell("B1").AsRange().Grow().RangeAddress.ToString(), Is.EqualTo("A1:C2"));

            Assert.That(ws.Cell("F5").AsRange().Grow().RangeAddress.ToString(), Is.EqualTo("E4:G6"));
            Assert.That(ws.Cell("F5").AsRange().Grow(2).RangeAddress.ToString(), Is.EqualTo("D3:H7"));
            Assert.That(ws.Cell("F5").AsRange().Grow(100).RangeAddress.ToString(), Is.EqualTo("A1:DB105"));
        }

        [Test]
        public void ShrinkRange()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A1").AsRange().Shrink(), Is.Null);
                Assert.That(ws.Range("B2:C3").Shrink(), Is.Null);
                Assert.That(ws.Range("B2:D4").Shrink().RangeAddress.ToString(), Is.EqualTo("C3:C3"));
                Assert.That(ws.Range("A1:Z26").Shrink(10).RangeAddress.ToString(), Is.EqualTo("K11:P16"));

                // Grow and shrink back
                Assert.That(ws.Cell("Z26").AsRange().Grow(10).Shrink(10).RangeAddress.ToString(), Is.EqualTo("Z26:Z26"));
            });
        }

        [Test]
        public void Intersection()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            Assert.Multiple(() =>
            {
                Assert.That(ws.Range("B9:I11").Intersection(ws.Range("D4:G16")).ToString(), Is.EqualTo("D9:G11"));
                Assert.That(ws.Range("E9:I11").Intersection(ws.Range("D4:G16")).ToString(), Is.EqualTo("E9:G11"));
                Assert.That(ws.Cell("E9").AsRange().Intersection(ws.Range("D4:G16")).ToString(), Is.EqualTo("E9:E9"));
                Assert.That(ws.Range("D4:G16").Intersection(ws.Cell("E9").AsRange()).ToString(), Is.EqualTo("E9:E9"));
            });

            XLRangeAddress rangeAddress;

            rangeAddress = (XLRangeAddress)ws.Cell("C3").AsRange().Intersection(ws.Cell("A1").AsRange());
            Assert.That(rangeAddress.IsValid, Is.False);

            rangeAddress = (XLRangeAddress)ws.Cell("A1").AsRange().Intersection(ws.Cell("C3").AsRange());
            Assert.That(rangeAddress.IsValid, Is.False);

            Assert.That(ws.Range("A1:C3").Intersection(null), Is.Null);

            var otherWs = wb.AddWorksheet("Sheet2");
            Assert.That(ws.Intersection(otherWs), Is.Null);
            Assert.That(ws.Cell("A1").AsRange().Intersection(otherWs.Cell("A2").AsRange()), Is.Null);
        }

        [Test]
        public void Union()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            Assert.Multiple(() =>
            {
                Assert.That(ws.Range("B9:I11").Union(ws.Range("D4:G16")).Count(), Is.EqualTo(64));
                Assert.That(ws.Range("E9:I11").Union(ws.Range("D4:G16")).Count(), Is.EqualTo(58));
                Assert.That(ws.Cell("E9").AsRange().Union(ws.Range("D4:G16")).Count(), Is.EqualTo(52));
                Assert.That(ws.Range("D4:G16").Union(ws.Cell("E9").AsRange()).Count(), Is.EqualTo(52));

                Assert.That(ws.Cell("A1").AsRange().Union(ws.Cell("C3").AsRange()).Count(), Is.EqualTo(2));

                Assert.That(ws.Range("A1:C3").Union(null).Count(), Is.EqualTo(9));
            });

            var otherWs = wb.AddWorksheet("Sheet2");
            Assert.That(ws.Union(otherWs).Any(), Is.False);
            Assert.That(ws.Cell("A1").AsRange().Union(otherWs.Cell("A2").AsRange()).Any(), Is.False);
        }

        [Test]
        public void Difference()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            Assert.Multiple(() =>
            {
                Assert.That(ws.Range("B9:I11").Difference(ws.Range("D4:G16")).Count(), Is.EqualTo(12));
                Assert.That(ws.Range("E9:I11").Difference(ws.Range("D4:G16")).Count(), Is.EqualTo(6));
                Assert.That(ws.Cell("E9").AsRange().Difference(ws.Range("D4:G16")).Count(), Is.EqualTo(0));
                Assert.That(ws.Range("D4:G16").Difference(ws.Cell("E9").AsRange()).Count(), Is.EqualTo(51));

                Assert.That(ws.Cell("A1").AsRange().Difference(ws.Cell("C3").AsRange()).Count(), Is.EqualTo(1));

                Assert.That(ws.Range("A1:C3").Difference(null).Count(), Is.EqualTo(9));
            });

            var otherWs = wb.AddWorksheet("Sheet2");
            Assert.That(ws.Difference(otherWs).Any(), Is.False);
            Assert.That(ws.Cell("A1").AsRange().Difference(otherWs.Cell("A2").AsRange()).Any(), Is.False);
        }

        [Test]
        public void SurroundingCells()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            Assert.Multiple(() =>
            {
                Assert.That(ws.FirstCell().AsRange().SurroundingCells().Count(), Is.EqualTo(3));
                Assert.That(ws.Cell("C3").AsRange().SurroundingCells().Count(), Is.EqualTo(8));
                Assert.That(ws.Range("C3:D6").AsRange().SurroundingCells().Count(), Is.EqualTo(16));

                Assert.That(ws.Range("C3:D6").AsRange().SurroundingCells(c => !c.IsEmpty()).Count(), Is.EqualTo(0));
            });
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeAbove1()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("B2:E3").Clear(XLClearOptions.ConditionalFormats);

            Assert.Multiple(() =>
            {
                Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
                Assert.That(ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative(), Is.EqualTo("C4:D7"));
            });
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeAbove2()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("C3:D3").Clear(XLClearOptions.ConditionalFormats);

            Assert.Multiple(() =>
            {
                Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
                Assert.That(ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative(), Is.EqualTo("C4:D7"));
            });
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeBelow1()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("B7:E8").Clear(XLClearOptions.ConditionalFormats);

            Assert.Multiple(() =>
            {
                Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
                Assert.That(ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative(), Is.EqualTo("C3:D6"));
            });
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeBelow2()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("C7:D7").Clear(XLClearOptions.ConditionalFormats);

            Assert.Multiple(() =>
            {
                Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
                Assert.That(ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative(), Is.EqualTo("C3:D6"));
            });
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeRowInMiddle()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("C5:E5").Clear(XLClearOptions.ConditionalFormats);

            Assert.Multiple(() =>
            {
                Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
                Assert.That(ws.ConditionalFormats.First().Ranges.First().RangeAddress.ToStringRelative(), Is.EqualTo("C3:D4"));
            });
            Assert.That(ws.ConditionalFormats.First().Ranges.Last().RangeAddress.ToStringRelative(), Is.EqualTo("C6:D7"));
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeColumnInMiddle()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:G4").AddConditionalFormat();
            ws.Range("E2:E4").Clear(XLClearOptions.ConditionalFormats);

            Assert.Multiple(() =>
            {
                Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
                Assert.That(ws.ConditionalFormats.First().Ranges.First().RangeAddress.ToStringRelative(), Is.EqualTo("C3:D4"));
            });
            Assert.That(ws.ConditionalFormats.First().Ranges.Last().RangeAddress.ToStringRelative(), Is.EqualTo("F3:G4"));
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeContainsFormatWhole()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:G4").AddConditionalFormat();
            ws.Range("B2:G4").Clear(XLClearOptions.ConditionalFormats);

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(0));
        }

        [Test]
        public void NoClearConditionalFormattingsWhenRangePartiallySuperimposed()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:G4").AddConditionalFormat();
            ws.Range("C2:D3").Clear(XLClearOptions.ConditionalFormats);

            Assert.Multiple(() =>
            {
                Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
                Assert.That(ws.ConditionalFormats.Single().Ranges, Has.Count.EqualTo(1));
            });
            Assert.That(ws.ConditionalFormats.Single().Ranges.Single().RangeAddress.ToStringRelative(), Is.EqualTo("C3:G4"));
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

            Assert.Multiple(() =>
            {
                Assert.That(ranges.Count, Is.EqualTo(0));
                // if ranges were not disposed they addresses should change
                Assert.That(rangesCopy.First().RangeAddress.ToString(), Is.EqualTo("B1:B2"));
            });
            Assert.That(rangesCopy.Last().RangeAddress.ToString(), Is.EqualTo("C1:C2"));
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

            Assert.That(ranges, Has.Count.EqualTo(1));
            Assert.That(ranges.Single().RangeAddress.ToString(), Is.EqualTo("A1:A2"));
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

            Assert.That(actualRanges, Has.Count.EqualTo(expectedRanges.Count));
            for (int i = 0; i < actualRanges.Count; i++)
            {
                Assert.That(actualRanges[i], Is.EqualTo(expectedRanges[i]));
            }
        }

        [Test]
        public void ClearRangeRemovesSparklines()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.SparklineGroups.Add("B1:B3", "C1:E3");

            ws.Range("B1:C1").Clear(XLClearOptions.All);
            ws.Range("B2:C2").Clear(XLClearOptions.Sparklines);

            Assert.Multiple(() =>
            {
                Assert.That(ws.SparklineGroups.Single().Count(), Is.EqualTo(1));
                Assert.That(ws.Cell("B1").HasSparkline, Is.False);
                Assert.That(ws.Cell("B2").HasSparkline, Is.False);
                Assert.That(ws.Cell("B3").HasSparkline, Is.True);
            });
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

            Assert.That(actualAddresses, Is.EqualTo(expectedResult));
        }

        [Test]
        public void Sorting_moves_values_and_fixes_formula_references()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            var range = ws.Cell("A1").InsertData(new object[]
            {
                ("Price", "Amount", "Sales"),
                (7, 5, Blank.Value),
                (2, 14, Blank.Value),
                (32, 2, Blank.Value),
                (6, 9, Blank.Value)
            });
            ws.Cell("C2").FormulaA1 = "A2*B2 & \"(Cake)\""; // 35
            ws.Cell("C3").FormulaA1 = "A3*B3 & \"(Pie)\""; // 28
            ws.Cell("C4").FormulaA1 = "A4*B4 & \"(Waffle)\""; // 64
            ws.Cell("C5").FormulaA1 = "A5*B5 & \"(Shortcake)\""; // 54

            // Sort uses cached values - update them
            ws.RecalculateAllFormulas();

            range.Sort("3 DESC");

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A2").Value, Is.EqualTo(32));
                Assert.That(ws.Cell("A3").Value, Is.EqualTo(6));
                Assert.That(ws.Cell("A4").Value, Is.EqualTo(7));
                Assert.That(ws.Cell("A5").Value, Is.EqualTo(2));

                Assert.That(ws.Cell("B2").Value, Is.EqualTo(2));
                Assert.That(ws.Cell("B3").Value, Is.EqualTo(9));
                Assert.That(ws.Cell("B4").Value, Is.EqualTo(5));
                Assert.That(ws.Cell("B5").Value, Is.EqualTo(14));

                // Formulas has been moved around and their coordinates fixed after move
                Assert.That(ws.Cell("C2").FormulaA1, Is.EqualTo("A2*B2 & \"(Waffle)\""));
                Assert.That(ws.Cell("C3").FormulaA1, Is.EqualTo("A3*B3 & \"(Shortcake)\""));
                Assert.That(ws.Cell("C4").FormulaA1, Is.EqualTo("A4*B4 & \"(Cake)\""));
                Assert.That(ws.Cell("C5").FormulaA1, Is.EqualTo("A5*B5 & \"(Pie)\""));
            });
        }

        [TestCase("PY(4)", "_xlfn._xlws.PY(4)")]
        [TestCase("2 + CHISQ.INV(0.6,2)", "2 + _xlfn.CHISQ.INV(0.6,2)")]
        [TestCase("2 + _xlfn.CHISQ.INV(0.6,2)", "2 + _xlfn.CHISQ.INV(0.6,2)")]
        public void FormulaArrayA1_adds_prefix_to_future_functions(string formula, string expected)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:B2").FormulaArrayA1 = formula;
            var masterCellFormula = ws.Cell("A1").FormulaA1;
            Assert.That(masterCellFormula, Is.EqualTo(expected));
        }
    }
}
