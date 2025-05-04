using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.Ranges
{
    [TestFixture]
    public class UsedAndUnusedCellsTests
    {
        private XLWorkbook workbook;

        [TearDown]
        public void Cleanup()
        {
            workbook.Dispose();
        }
        
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
            ws.Cell(6, 2).Style.Fill.BackgroundColor = XLColor.Red;
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
            Assert.That(i, Is.EqualTo(2));

            i = 0;
            row = workbook.Worksheets.First().FirstRow().RowBelow();
            foreach (var cell in row.Cells())
            {
                i++;
            }
            Assert.That(i, Is.EqualTo(1));

            i = 0;
            row = workbook.Worksheets.First().LastRowUsed(XLCellsUsedOptions.All);
            Assert.That(row.RowNumber(), Is.EqualTo(6));
            foreach (var cell in row.Cells())
            {
                i++;
            }
            Assert.That(i, Is.EqualTo(1));

            i = 0;
            row = workbook.Worksheets.First().LastRowUsed(XLCellsUsedOptions.All);
            Assert.That(row.RowNumber(), Is.EqualTo(6));
            foreach (var cell in row.CellsUsed())
            {
                i++;
            }
            Assert.That(i, Is.EqualTo(0));
        }

        [Test(Description = "See 1443")]
        public void FirstRowUsedRegression()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Range("B3:F6").SetValue(100);

            Assert.That(ws.FirstRowUsed(XLCellsUsedOptions.AllContents).RowNumber(), Is.EqualTo(3));
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
            Assert.That(i, Is.EqualTo(3));

            i = 0;
            row = workbook.Worksheets.First().FirstRow().RowBelow(); //This row has no empty cells BETWEEN used cells
            foreach (var cell in row.Cells(false))
            {
                i++;
            }
            Assert.That(i, Is.EqualTo(1));
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
            Assert.That(i, Is.EqualTo(2));

            i = 0;
            column = workbook.Worksheets.First().FirstColumn().ColumnRight().ColumnRight();
            foreach (var cell in column.Cells())
            {
                i++;
            }
            Assert.That(i, Is.EqualTo(1));

            i = 0;
            column = workbook.Worksheets.First().Column(2);
            foreach (var cell in column.Cells())
            {
                i++;
            }
            Assert.That(i, Is.EqualTo(3));

            i = 0;
            column = workbook.Worksheets.First().Column(2);
            foreach (var cell in column.CellsUsed())
            {
                i++;
            }
            Assert.That(i, Is.EqualTo(2));
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
            Assert.That(i, Is.EqualTo(4));

            i = 0;
            column = workbook.Worksheets.First().FirstColumn().ColumnRight().ColumnRight(); //This column has no empty cells BETWEEN used cells
            foreach (var cell in column.Cells(false))
            {
                i++;
            }
            Assert.That(i, Is.EqualTo(1));
        }

        [Test]
        public void CountCellsInWorksheet()
        {
            var ws = workbook.Worksheets.First();
            int i = 0;

            foreach (var cell in ws.Cells()) // All cells with content or formats
            {
                i++;
            }
            Assert.That(i, Is.EqualTo(6));
        }

        [Test]
        public void CountUsedCellsInWorksheet()
        {
            var ws = workbook.Worksheets.First();
            int i = 0;

            foreach (var cell in ws.CellsUsed()) // Only used cells in worksheet
            {
                i++;
            }
            Assert.That(i, Is.EqualTo(5));
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
            Assert.That(i, Is.EqualTo(18));
        }

        [Test]
        public void GetCellsUsedNonRectangular()
        {
            using XLWorkbook wb = new XLWorkbook();
            var sheet = wb.AddWorksheet("page1");

            sheet.Range("C1:E1").Value = "row1";
            sheet.Range("A2:E2").Value = "row2";

            var used = sheet.RangeUsed().RangeAddress.ToString(XLReferenceStyle.A1);

            Assert.That(used, Is.EqualTo("A1:E2"));
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
            using XLWorkbook wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            foreach (var cellAddress in cellsWithValues)
            {
                ws.Cell(cellAddress).Value = "Not empty";
            }
            ws.Range("B2:D2").Merge();

            var options = includeFormatting
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents | XLCellsUsedOptions.MergedRanges;
            var actual = ws.RangeUsed(options).RangeAddress;

            Assert.That(actual.ToString(), Is.EqualTo(expectedRange));
        }

        [Test]
        public void LastCellUsedPredicateConsidersMergedRanges()
        {
            using XLWorkbook wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.Red;
            ws.Cell("A2").Style.Fill.BackgroundColor = XLColor.Yellow;
            ws.Cell("A3").Style.Fill.BackgroundColor = XLColor.Green;
            ws.Range("A1:C1").Merge();
            ws.Range("A2:C2").Merge();
            ws.Range("A3:C3").Merge();

            var actual = ws.LastCellUsed(XLCellsUsedOptions.All,
                c => c.Style.Fill.BackgroundColor == XLColor.Yellow);

            Assert.That(actual.Address.ToString(), Is.EqualTo("C2"));
        }

        [Test]
        public void FirstCellUsedPredicateConsidersMergedRanges()
        {
            using XLWorkbook wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.Red;
            ws.Cell("A2").Style.Fill.BackgroundColor = XLColor.Yellow;
            ws.Cell("A3").Style.Fill.BackgroundColor = XLColor.Green;
            ws.Range("A1:C1").Merge();
            ws.Range("A2:C2").Merge();
            ws.Range("A3:C3").Merge();

            var actual = ws.FirstCellUsed(XLCellsUsedOptions.All,
                c => c.Style.Fill.BackgroundColor == XLColor.Yellow);

            Assert.That(actual.Address.ToString(), Is.EqualTo("A2"));
        }

        [Test]
        public void ApplyingDataValidationMakesCellNotEmpty()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Range("B2:B12").CreateDataValidation()
                .Decimal.EqualOrGreaterThan(0);

            var usedCells = ws.CellsUsed(XLCellsUsedOptions.All).ToList();

            Assert.That(usedCells, Has.Count.EqualTo(11));
            Assert.Multiple(() =>
            {
                Assert.That(usedCells.First().Address.ToString(), Is.EqualTo("B2"));
                Assert.That(usedCells.Last().Address.ToString(), Is.EqualTo("B12"));
            });
        }

        [Test]
        public void MergeMakesCellNotEmpty()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Range("B2:B12").Merge();

            var usedCells = ws.CellsUsed(XLCellsUsedOptions.All).ToList();

            Assert.That(usedCells, Has.Count.EqualTo(11));
            Assert.Multiple(() =>
            {
                Assert.That(usedCells.First().Address.ToString(), Is.EqualTo("B2"));
                Assert.That(usedCells.Last().Address.ToString(), Is.EqualTo("B12"));
            });
        }

        [Test]
        public void FirstCellUsedNotHangingOnLargeCFRules()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.AddConditionalFormat().WhenIsBlank().Fill.SetBackgroundColor(XLColor.Gold);

            var firstCell = ws.FirstCellUsed(XLCellsUsedOptions.All);

            Assert.Multiple(() =>
            {
                Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count(), Is.EqualTo(0));
                Assert.That(firstCell.Address.ToString(), Is.EqualTo("A1"));
            });
        }

        [Test]
        public void LastCellUsedNotHangingOnLargeCFRules()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.AddConditionalFormat().WhenIsBlank().Fill.SetBackgroundColor(XLColor.Gold);

            var lastCell = ws.LastCellUsed(XLCellsUsedOptions.All);

            Assert.Multiple(() =>
            {
                Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count(), Is.EqualTo(0));
                Assert.That(lastCell.Address.ToString(), Is.EqualTo(XLHelper.LastCell));
            });
        }

        [Test]
        public void FirstCellUsedNotHangingOnLargeDVRules()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.CreateDataValidation().WholeNumber.GreaterThan(0);

            var firstCell = ws.FirstCellUsed(XLCellsUsedOptions.All);

            Assert.Multiple(() =>
            {
                Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count(), Is.EqualTo(0));
                Assert.That(firstCell.Address.ToString(), Is.EqualTo("A1"));
            });
        }

        [Test]
        public void LastCellUsedNotHangingOnLargeDVRules()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.CreateDataValidation().WholeNumber.GreaterThan(0);

            var lastCell = ws.LastCellUsed(XLCellsUsedOptions.All);

            Assert.Multiple(() =>
            {
                Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count(), Is.EqualTo(0));
                Assert.That(lastCell.Address.ToString(), Is.EqualTo(XLHelper.LastCell));
            });
        }

        [Test]
        public void FirstCellUsedNotHangingOnLargeMergedRanges()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Merge();

            var firstCell = ws.FirstCellUsed(XLCellsUsedOptions.All);

            Assert.Multiple(() =>
            {
                Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count(), Is.EqualTo(0));
                Assert.That(firstCell.Address.ToString(), Is.EqualTo("A1"));
            });
        }

        [Test]
        public void LastCellUsedNotHangingOnLargeMergedRanges()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Merge();

            var lastCell = ws.LastCellUsed(XLCellsUsedOptions.All);

            Assert.Multiple(() =>
            {
                Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count(), Is.EqualTo(0));
                Assert.That(lastCell.Address.ToString(), Is.EqualTo(XLHelper.LastCell));
            });
        }
    }
}
