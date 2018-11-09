using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML_Tests.Excel.Ranges
{
    [TestFixture]
    public class UsedAndUnusedCellsTests
    {
        private XLWorkbook workbook;

        [SetUp]
        public void SetupWorkbook()
        {
            workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");
            ws.Cell(1, 1).Value = "A1";
            ws.Cell(1, 3).Value = "C1";
            ws.Cell(2, 2).Value = "B2";
            ws.Cell(4, 1).Value = "A4";
            ws.Cell(5, 2).Value = "B5";
        }

        [Test]
        public void CountUsedCellsInRow()
        {
            int i = 0;
            var row = workbook.Worksheets.First().FirstRow();
            foreach (var cell in row.Cells()) // Cells() returns UnUsed cells by default
            {
                i++;
            }
            Assert.AreEqual(2, i);

            i = 0;
            row = workbook.Worksheets.First().FirstRow().RowBelow();
            foreach (var cell in row.Cells())
            {
                i++;
            }
            Assert.AreEqual(1, i);
        }

        [Test]
        public void CountAllCellsInRow()
        {
            int i = 0;
            var row = workbook.Worksheets.First().FirstRow();
            foreach (var cell in row.Cells(false)) // All cells in range between first and last cells used
            {
                i++;
            }
            Assert.AreEqual(3, i);

            i = 0;
            row = workbook.Worksheets.First().FirstRow().RowBelow(); //This row has no empty cells BETWEEN used cells
            foreach (var cell in row.Cells(false))
            {
                i++;
            }
            Assert.AreEqual(1, i);
        }

        [Test]
        public void CountUsedCellsInColumn()
        {
            int i = 0;
            var column = workbook.Worksheets.First().FirstColumn();
            foreach (var cell in column.Cells()) // Cells() returns UnUsed cells by default
            {
                i++;
            }
            Assert.AreEqual(2, i);

            i = 0;
            column = workbook.Worksheets.First().FirstColumn().ColumnRight().ColumnRight();
            foreach (var cell in column.Cells())
            {
                i++;
            }
            Assert.AreEqual(1, i);
        }

        [Test]
        public void CountAllCellsInColumn()
        {
            int i = 0;
            var column = workbook.Worksheets.First().FirstColumn();
            foreach (var cell in column.Cells(false)) // All cells in range between first and last cells used
            {
                i++;
            }
            Assert.AreEqual(4, i);

            i = 0;
            column = workbook.Worksheets.First().FirstColumn().ColumnRight().ColumnRight(); //This column has no empty cells BETWEEN used cells
            foreach (var cell in column.Cells(false))
            {
                i++;
            }
            Assert.AreEqual(1, i);
        }

        [Test]
        public void CountUsedCellsInWorksheet()
        {
            var ws = workbook.Worksheets.First();
            int i = 0;

            foreach (var cell in ws.Cells()) // Only used cells in worksheet
            {
                i++;
            }
            Assert.AreEqual(5, i);
        }

        [Test]
        public void CountAllCellsInWorksheet()
        {
            var ws = workbook.Worksheets.First();
            int i = 0;

            foreach (var cell in ws.Cells(false)) // All cells in range between first and last cells used (cartesian product of range)
            {
                i++;
            }
            Assert.AreEqual(15, i);
        }

        [Test]
        public void GetCellsUsedNonRectangular()
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                var sheet = wb.AddWorksheet("page1");

                sheet.Range("C1:E1").Value = "row1";
                sheet.Range("A2:E2").Value = "row2";

                var used = sheet.RangeUsed().RangeAddress.ToString(XLReferenceStyle.A1);

                Assert.AreEqual("A1:E2", used);
            }
        }

        [TestCase(true, "A1:D2", "A1")]
        [TestCase(true, "A2:D2", "A2")]
        [TestCase(true, "A1:D2", "A1", "B2")]
        [TestCase(true, "B2:D3", "C3")]
        [TestCase(true, "B2:F4", "F4")]
        [TestCase(false, "A1:D2", "A1")]
        [TestCase(false, "A2:D2", "A2")]
        [TestCase(false, "A1:D2", "A1", "B2")]
        [TestCase(false, "B2:D3", "C3")]
        [TestCase(false, "B2:F4", "F4")]
        public void RangeUsedIncludesMergedCells(bool includeFormatting, string expectedRange,
            params string[] cellsWithValues)
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                foreach (var cellAddress in cellsWithValues)
                {
                    ws.Cell(cellAddress).Value = "Not empty";
                }
                ws.Range("B2:D2").Merge();

                var options = includeFormatting
                    ? XLCellsUsedOptions.All
                    : XLCellsUsedOptions.AllContents;
                var actual = ws.RangeUsed(options).RangeAddress;

                Assert.AreEqual(expectedRange, actual.ToString());
            }
        }

        [Test]
        public void LastCellUsedPredicateConsidersMergedRanges()
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.Red;
                ws.Cell("A2").Style.Fill.BackgroundColor = XLColor.Yellow;
                ws.Cell("A3").Style.Fill.BackgroundColor = XLColor.Green;
                ws.Range("A1:C1").Merge();
                ws.Range("A2:C2").Merge();
                ws.Range("A3:C3").Merge();

                var actual = ws.LastCellUsed(XLCellsUsedOptions.All,
                    c => c.Style.Fill.BackgroundColor == XLColor.Yellow);

                Assert.AreEqual("C2", actual.Address.ToString());
            }
        }

        [Test]
        public void FirstCellUsedPredicateConsidersMergedRanges()
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.Red;
                ws.Cell("A2").Style.Fill.BackgroundColor = XLColor.Yellow;
                ws.Cell("A3").Style.Fill.BackgroundColor = XLColor.Green;
                ws.Range("A1:C1").Merge();
                ws.Range("A2:C2").Merge();
                ws.Range("A3:C3").Merge();

                var actual = ws.FirstCellUsed(XLCellsUsedOptions.All,
                    c => c.Style.Fill.BackgroundColor == XLColor.Yellow);

                Assert.AreEqual("A2", actual.Address.ToString());
            }
        }

        [Test]
        public void ApplyingDataValidationMakesCellNotEmpty()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                ws.Range("B2:B12").SetDataValidation()
                    .Decimal.EqualOrGreaterThan(0);

                var usedCells = ws.CellsUsed(XLCellsUsedOptions.All).ToList();

                Assert.AreEqual(11, usedCells.Count);
                Assert.AreEqual("B2", usedCells.First().Address.ToString());
                Assert.AreEqual("B12", usedCells.Last().Address.ToString());
            }
        }

        [Test]
        public void MergeMakesCellNotEmpty()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                ws.Range("B2:B12").Merge();

                var usedCells = ws.CellsUsed(XLCellsUsedOptions.All).ToList();

                Assert.AreEqual(11, usedCells.Count);
                Assert.AreEqual("B2", usedCells.First().Address.ToString());
                Assert.AreEqual("B12", usedCells.Last().Address.ToString());
            }
        }

        [Test]
        public void FirstCellUsedNotHangingOnLargeCFRules()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                ws.AddConditionalFormat().WhenIsBlank().Fill.SetBackgroundColor(XLColor.Gold);

                var firstCell = ws.FirstCellUsed(XLCellsUsedOptions.All);

                Assert.AreEqual(1, (ws as XLWorksheet).Internals.CellsCollection.Count);
                Assert.AreEqual("A1", firstCell.Address.ToString());
            }
        }

        [Test]
        public void LastCellUsedNotHangingOnLargeCFRules()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                ws.AddConditionalFormat().WhenIsBlank().Fill.SetBackgroundColor(XLColor.Gold);

                var lastCell = ws.LastCellUsed(XLCellsUsedOptions.All);

                Assert.AreEqual(1, (ws as XLWorksheet).Internals.CellsCollection.Count);
                Assert.AreEqual(XLHelper.LastCell, lastCell.Address.ToString());
            }
        }

        [Test]
        public void FirstCellUsedNotHangingOnLargeDVRules()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                ws.SetDataValidation().WholeNumber.GreaterThan(0);

                var firstCell = ws.FirstCellUsed(XLCellsUsedOptions.All);

                Assert.AreEqual(1, (ws as XLWorksheet).Internals.CellsCollection.Count);
                Assert.AreEqual("A1", firstCell.Address.ToString());
            }
        }

        [Test]
        public void LastCellUsedNotHangingOnLargeDVRules()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                ws.SetDataValidation().WholeNumber.GreaterThan(0);

                var lastCell = ws.LastCellUsed(XLCellsUsedOptions.All);

                Assert.AreEqual(1, (ws as XLWorksheet).Internals.CellsCollection.Count);
                Assert.AreEqual(XLHelper.LastCell, lastCell.Address.ToString());
            }
        }

        [Test]
        public void FirstCellUsedNotHangingOnLargeMergedRanges()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                ws.Merge();

                var firstCell = ws.FirstCellUsed(XLCellsUsedOptions.All);

                Assert.AreEqual(1, (ws as XLWorksheet).Internals.CellsCollection.Count);
                Assert.AreEqual("A1", firstCell.Address.ToString());
            }
        }

        [Test]
        public void LastCellUsedNotHangingOnLargeMergedRanges()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                ws.Merge();

                var lastCell = ws.LastCellUsed(XLCellsUsedOptions.All);

                Assert.AreEqual(2, (ws as XLWorksheet).Internals.CellsCollection.Count);
                Assert.AreEqual(XLHelper.LastCell, lastCell.Address.ToString());
            }
        }

    }
}
