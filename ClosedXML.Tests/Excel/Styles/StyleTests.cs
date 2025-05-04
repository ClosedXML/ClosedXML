using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class StyleTests
    {
        [Test]
        public void EmptyCellWithQuotePrefixNotTreatedAsEmpty()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.FirstCell().SetValue("Empty cell with quote prefix:");
                var cell = ws.FirstCell().CellRight() as XLCell;

                Assert.That(cell.IsEmpty(), Is.True);
                cell.Value = string.Empty;
                cell.Style.IncludeQuotePrefix = true;

                Assert.Multiple(() =>
                {
                    Assert.That(cell.IsEmpty(), Is.True);
                    Assert.That(cell.IsEmpty(XLCellsUsedOptions.All), Is.False);
                });

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                var cell = (XLCell)ws.Cell("B1");
                Assert.Multiple(() =>
                {
                    Assert.That(cell.MemorySstId, Is.EqualTo(1));

                    Assert.That(cell.IsEmpty(), Is.True);
                    Assert.That(cell.IsEmpty(XLCellsUsedOptions.All), Is.False);
                });
            }
        }

        [TestCase("A1", TestName = "First cell")]
        [TestCase("A2", TestName = "Cell from initialized row")]
        [TestCase("B1", TestName = "Cell from initialized column")]
        [TestCase("D4", TestName = "Initialized cell")]
        [TestCase("F6", TestName = "Non-initialized cell")]
        public void CellTakesWorksheetStyle(string cellAddress)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Column(2);
            ws.Row(2);
            ws.Cell("D4").Value = "Non empty";
            ws.Style.Font.SetFontName("Arial");
            ws.Style.Font.SetFontSize(9);

            var cell = ws.Cell(cellAddress);
            Assert.Multiple(() =>
            {
                Assert.That(cell.Style.Font.FontName, Is.EqualTo("Arial"));
                Assert.That(cell.Style.Font.FontSize, Is.EqualTo(9));
            });
        }

        [TestCaseSource(nameof(StylizedEntities))]
        public void WorksheetStyleAffectsAllNestedEntities(Func<IXLWorksheet, IXLStyle> getEntityStyle)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Style.Font.FontSize = 8;

            var style = getEntityStyle(ws);

            Assert.That(style.Font.FontSize, Is.EqualTo(8));
        }

        // https://github.com/ClosedXML/ClosedXML/issues/1813
        [Test]
        public void RowColors()
        {
            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook();
                {
                    var ws = wb.Worksheets.Add("Row Settings 1");
                    ws.Style.Fill.BackgroundColor = XLColor.Green;

                    var row1 = ws.Row(2);
                    row1.Style.Fill.BackgroundColor = XLColor.Red;
                    row1.Height = 30;

                    var row2 = ws.Row(4);
                    row2.Style.Fill.BackgroundColor = XLColor.DarkOrange;
                    row2.Height = 3;
                }

                {
                    var ws = wb.Worksheets.Add("Row Settings 2");
                    ws.Style.Fill.BackgroundColor = XLColor.Red;

                    var row1 = ws.Row(2);
                    row1.Style.Fill.BackgroundColor = XLColor.Red;

                    var row2 = ws.Row(4);
                    row2.Style.Fill.BackgroundColor = XLColor.DarkOrange;
                    row2.Height = 3;
                }

                {
                    var ws = wb.Worksheets.Add("Row Settings 3");
                    ws.Style.Fill.BackgroundColor = XLColor.Red;

                    var row1 = ws.Row(2);
                    row1.Style.Fill.BackgroundColor = XLColor.Red;
                    row1.Height = 30;

                    var row2 = ws.Row(4);
                    row2.Style.Fill.BackgroundColor = XLColor.DarkOrange;
                    row2.Height = 3;
                }

                return wb;
            }, @"Other\StyleReferenceFiles\RowColors\output.xlsx");
        }

        [Test]
        public void Style_for_cells_without_explicitly_set_style_uses_combination_of_row_and_columns_styles()
        {
            // If a style for a cell hasn't been explicitly set (e.g. though `cell.Style.Font
            // .SetBold(true)`), it is not yet instantiated to save memory and the actual value
            // is determined by the column style and row style. Generally speaking, the axis that
            // had its value set explicitly has a precedence, but because we can't detect that with
            // current structure, use difference from worksheet as an indication of explicitly set
            // value instead.
            // If row and column style components differ, the cells at the cross are pinged, thus test
            // sets different components for each axis.
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            var rowStyle = ws.Row(4).Style
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                .Fill.SetBackgroundColor(XLColor.Blue)
                .SetIncludeQuotePrefix()
                .Protection.SetLocked(true);

            var colStyle = ws.Column(2).Style
                .Border.SetBottomBorder(XLBorderStyleValues.Double)
                .Font.SetFontName("Arial")
                .NumberFormat.SetNumberFormatId((int)XLPredefinedFormat.Number.Precision2);

            var crossCellStyle = ws.Cell(4, 2).Style;
            Assert.Multiple(() =>
            {
                Assert.That(crossCellStyle.Alignment.Horizontal, Is.EqualTo(XLAlignmentHorizontalValues.Center));
                Assert.That(crossCellStyle.Border.BottomBorder, Is.EqualTo(XLBorderStyleValues.Double));
                Assert.That(crossCellStyle.Fill.BackgroundColor, Is.EqualTo(XLColor.Blue));
                Assert.That(crossCellStyle.IncludeQuotePrefix, Is.True);
                Assert.That(crossCellStyle.NumberFormat.NumberFormatId, Is.EqualTo((int)XLPredefinedFormat.Number.Precision2));
                Assert.That(crossCellStyle.Protection.Locked, Is.True);
            });

            var rowCellStyle = ws.Cell(4, 3).Style;
            Assert.That(rowCellStyle, Is.EqualTo(rowStyle));

            var colCellStyle = ws.Cell(5, 2).Style;
            Assert.That(colCellStyle, Is.EqualTo(colStyle));
        }

        private static IEnumerable<TestCaseData> StylizedEntities
        {
            get
            {
                var t = nameof(WorksheetStyleAffectsAllNestedEntities);
                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Style)).SetName(t + ": Worksheet");

                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Columns().Style)).SetName(t + ": Columns()");
                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Columns(1, 3).Style)).SetName(t + ": Columns(1, 3)");
                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Columns("B:F").Style)).SetName(t + ": Columns(\"B:F\")");
                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Columns("B", "F").Style)).SetName(t + ": Columns(\"B\", \"F\")");
                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Column(5).Style)).SetName(t + ": Column(5)");
                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Column("D").Style)).SetName(t + ": Column(\"D\")");

                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Rows().Style)).SetName(t + ": Rows()");
                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Rows(1, 3).Style)).SetName(t + ": Rows(1, 3)");
                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Rows("1:3").Style)).SetName(t + ": Rows(\"1:3\")");
                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Row(5).Style)).SetName(t + ": Row(5)");

                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Cells().Style)).SetName(t + ": Cells()");
                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Cells("B2,D4").Style)).SetName(t + ": Cells(\"B2, D4\")");
                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Cell("F6").Style)).SetName(t + ": Cell(\"F6\")");
                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Cell(2, 3).Style)).SetName(t + ": Cell(2, 3)");

                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Ranges("F6:H9,I8:K10").Style)).SetName(t + ": Ranges(\"F6:H9,I8:K10\")");
                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Range("G8:H10").Style)).SetName(t + ": Range(\"G8:H10\")");
                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Range("G8:H10").Column(1).Style)).SetName(t + ": Range(\"G8:H10\").Column(1)");
                yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>((ws) => ws.Range("G8:H10").Row(2).Style)).SetName(t + ": Range(\"G8:H10\").Row(2)");
            }
        }
    }
}
