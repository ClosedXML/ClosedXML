using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using NUnit.Framework;
using System.IO;
using System.Linq;

namespace ClosedXML_Tests.Excel
{
    [TestFixture]
    public class StyleTests
    {
        [Test]
        public void EmptyCellWithQuotePrefixNotTreatedAsEmpty()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet("Sheet1");
                    ws.FirstCell().SetValue("Empty cell with quote prefix:");
                    var cell = ws.FirstCell().CellRight() as XLCell;

                    Assert.IsTrue(cell.IsEmpty());
                    cell.Style.IncludeQuotePrefix = true;

                    Assert.IsTrue(cell.IsEmpty());
                    Assert.IsFalse(cell.IsEmpty(XLCellsUsedOptions.All));

                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheets.First();
                    var cell = ws.FirstCell().CellRight() as XLCell;
                    Assert.AreEqual(1, cell.SharedStringId);

                    Assert.IsTrue(cell.IsEmpty());
                    Assert.IsFalse(cell.IsEmpty(XLCellsUsedOptions.All));
                }
            }
        }

        [TestCase("A1", TestName = "First cell")]
        [TestCase("A2", TestName = "Cell from initialized row")]
        [TestCase("B1", TestName = "Cell from initialized column")]
        [TestCase("D4", TestName = "Initialized cell")]
        [TestCase("F6", TestName = "Non-initialized cell")]
        public void CellTakesWorksheetStyle(string cellAddress)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Column(2);
                ws.Row(2);
                ws.Cell("D4").Value = "Non empty";
                ws.Style.Font.SetFontName("Arial");
                ws.Style.Font.SetFontSize(9);

                var cell = ws.Cell(cellAddress);
                Assert.AreEqual("Arial", cell.Style.Font.FontName);
                Assert.AreEqual(9, cell.Style.Font.FontSize);
            }
        }

        [TestCaseSource(nameof(StylizedEntities))]
        public void WorksheetStyleAffectsAllNestedEntities(Func<IXLWorksheet, IXLStyle> getEntityStyle)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();

                ws.Style.Font.FontSize = 8;

                var style = getEntityStyle(ws);

                Assert.AreEqual(8, style.Font.FontSize);
            }
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

        [TestCase("A1", "Normal")]
        [TestCase("A2", "Good")]
        [TestCase("A3", "Bad")]
        [TestCase("B1", "Better")]
        [TestCase("B2", "Worse")]
        [TestCase("B3", "Normalish")]
        public void CanReadStyleNames(string cellAddress, string expectedName)
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\NamedStyles\input.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();

                var actualName = ws.Cell(cellAddress).Style.Name;

                Assert.AreEqual(expectedName, actualName);
            }
        }

        [Test]
        public void CanChangeStyleName()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var cell = ws.FirstCell();
                cell.Style.Name = "Test Style";

                Assert.AreEqual("Test Style", cell.Style.Name);
                Assert.AreSame(wb.NamedStyles["Test Style"], (cell.Style as XLStyle).Value);
            }
        }

        [Test]
        public void CannotChangeStyleNameToExisting()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var cell1 = ws.Cell("A1");
                var cell2 = ws.Cell("A2");

                cell1.Style.Name = "Style 1";

                Assert.Throws<InvalidOperationException>(() => cell2.Style.Name = "Style 1");
            }
        }

        [Test]
        public void StyleNamesAreCaseInsensitive()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var cell1 = ws.Cell("A1");
                var cell2 = ws.Cell("A2");

                cell1.Style.Name = "Style 1";

                Assert.Throws<InvalidOperationException>(() => cell2.Style.Name = "STYLE 1");
            }
        }

        [Test]
        public void CannotChangeNameOfDetachedStyle()
        {
            var style = new XLStyle(new XLStylizedEmpty(XLStyle.Default));

            Assert.Throws<InvalidOperationException>(() => style.Name = "Custom style");
        }

        [Test]
        public void CanCopyNamedStyleToAnotherWorkbook()
        {
            using (var wb1 = new XLWorkbook())
            using (var wb2 = new XLWorkbook())
            {
                var ws1 = wb1.AddWorksheet();
                var cell = ws1.FirstCell();
                cell.Style.Name = "Test Style";

                var ws2 = ws1.CopyTo(wb2, "Copy");

                Assert.AreEqual("Test Style", ws2.FirstCell().Style.Name);
                Assert.AreSame((ws2.FirstCell().Style as XLStyle).Value, wb2.NamedStyles["Test Style"]);
            }
        }

        [Test]
        public void CopiedStylesAreRenamedIfAlreadyExistsAndDiffers()
        {
            using (var wb1 = new XLWorkbook())
            using (var wb2 = new XLWorkbook())
            {
                var ws1 = wb1.AddWorksheet();
                var cell1 = ws1.FirstCell();
                cell1.Style.Name = "Test Style";

                var ws2 = wb2.AddWorksheet();
                var cell2 = ws2.FirstCell();
                cell2.Style.Fill.BackgroundColor = XLColor.Amber;
                cell2.Style.Name = "Test Style";

                var cell3 = ws2.Cell("B1");
                cell3.CopyFrom(cell1);

                Assert.AreEqual("Test Style 1", cell3.Style.Name);
            }
        }

        [Test]
        public void CopiedStylesAreRenamedIfAlreadyExistsAndEquals()
        {
            using (var wb1 = new XLWorkbook())
            using (var wb2 = new XLWorkbook())
            {
                var ws1 = wb1.AddWorksheet();
                var cell1 = ws1.FirstCell();
                cell1.Style.Fill.BackgroundColor = XLColor.Amber;
                cell1.Style.Name = "Test Style";

                var ws2 = wb2.AddWorksheet();
                var cell2 = ws2.FirstCell();
                cell2.Style.Fill.BackgroundColor = XLColor.Amber;
                cell2.Style.Name = "Test Style";

                var cell3 = ws2.Cell("B1");
                cell3.CopyFrom(cell1);

                Assert.AreEqual("Test Style", cell3.Style.Name);
            }
        }
        
        [Test]
        public void CanSaveStyleName()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet("Sheet1");
                    var cell = ws.FirstCell();
                    cell.Style.Fill.BackgroundColor = XLColor.Amber;
                    cell.Style.SetName("TestStyle");
                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheets.First();
                    var cell = ws.FirstCell();
                    Assert.AreEqual("TestStyle", cell.Style.Name);
                }
            }
        }

        [Test]
        public void CanLoadAndSaveNamedStyles()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\NamedStyles\input.xlsx")))
            using (var ms = new MemoryStream())
            {
                TestHelper.CreateAndCompare(() =>
                {
                    var wb = new XLWorkbook(stream);
                    wb.SaveAs(ms);
                    return wb;
                }, @"Other\StyleReferenceFiles\NamedStyles\output.xlsx");
            }
        }
    }
}
