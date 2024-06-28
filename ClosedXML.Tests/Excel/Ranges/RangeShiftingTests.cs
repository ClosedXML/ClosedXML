using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Ranges
{
    public class RangeShiftingTests
    {
        [Test]
        public void CellsContentShiftedAfterColumnDeleted()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                SetContent(ws.Cell("D4"));

                ws.Column("C").Delete();

                AssertContent(ws.Cell("C4"), "D4");
            }
        }

        [Test]
        public void CellsContentShiftedAfterRowDeleted()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                SetContent(ws.Cell("D4"));

                ws.Row(3).Delete();

                AssertContent(ws.Cell("D3"), "D4");
            }
        }

        [Test]
        public void CellsContentShiftedAfterColumnInserted()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                SetContent(ws.Cell("D4"));

                ws.Column("C").InsertColumnsBefore(1);

                AssertContent(ws.Cell("E4"), "D4");
            }
        }

        [Test]
        public void CellsContentShiftedAfterRowInserted()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                SetContent(ws.Cell("D4"));

                ws.Row(3).InsertRowsAbove(1);

                AssertContent(ws.Cell("D5"), "D4");
            }
        }

        [Test]
        public void CellsContentShiftAfterRangeDeleted()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                SetContent(ws.Cell("D4"));
                SetContent(ws.Cell("F8"));

                ws.Range("B2:C5").Delete(XLShiftDeletedCells.ShiftCellsLeft);
                ws.Range("E5:F7").Delete(XLShiftDeletedCells.ShiftCellsUp);

                AssertContent(ws.Cell("B4"), "D4");
                AssertContent(ws.Cell("F5"), "F8");
            }
        }

        [Theory]
        [TestCase("A5:F5")]
        [TestCase("A5:F6")]
        public void RangesBelowStayMergedAfterRangeDeleted(string deletedRangeAddress)
        {
            //There is an edge case when a merged range of same size as the deleted range got unmerged (see #2358)
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var deletedRange = ws.Range(deletedRangeAddress);
            var rangeHeight = deletedRange.LastRow().RowNumber() - deletedRange.FirstRow().RowNumber() + 1;
            var mergedRange = ws.Range(
                deletedRange.LastRow().RowNumber() + 1,
                deletedRange.FirstColumn().ColumnNumber(),
                deletedRange.LastRow().RowNumber() + rangeHeight,
                deletedRange.LastColumn().ColumnNumber()
            );
            mergedRange.Merge();

            deletedRange.Delete(XLShiftDeletedCells.ShiftCellsUp);

            Assert.IsTrue(mergedRange.IsMerged());
            Assert.AreEqual(deletedRangeAddress, mergedRange.RangeAddress.ToString());
        }

        [Theory]
        [TestCase("A5:A8")]
        [TestCase("A5:B8")]
        public void RangesToTheRightStayMergedAfterRangeDeleted(string deletedRangeAddress)
        {
            //There is an edge case when a merged range of same size as the deleted range got unmerged (see #2358)
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var deletedRange = ws.Range(deletedRangeAddress);
            var rangeWidth = deletedRange.LastColumn().ColumnNumber() - deletedRange.FirstColumn().ColumnNumber() + 1;
            var mergedRange = ws.Range(
                deletedRange.FirstRow().RowNumber(),
                deletedRange.LastColumn().ColumnNumber() + 1,
                deletedRange.LastRow().RowNumber(),
                deletedRange.LastColumn().ColumnNumber() + rangeWidth
            );
            mergedRange.Merge();

            deletedRange.Delete(XLShiftDeletedCells.ShiftCellsLeft);

            Assert.IsTrue(mergedRange.IsMerged());
            Assert.AreEqual(deletedRangeAddress, mergedRange.RangeAddress.ToString());
        }

        private void SetContent(IXLCell cell)
        {
            cell.FormulaA1 = $"\"Formula \" & \"{cell.Address}\"";
            cell.Style.Fill.SetBackgroundColor(XLColor.Green);
            cell.CreateComment().AddText("Some comment " + cell.Address);
        }

        private void AssertContent(IXLCell cell, string originalAddress)
        {
            Assert.AreEqual($"\"Formula \" & \"{originalAddress}\"", cell.FormulaA1);
            Assert.AreEqual(XLColor.Green, cell.Style.Fill.BackgroundColor);
            Assert.True(cell.HasComment);
            Assert.AreEqual($"Some comment {originalAddress}", cell.GetComment().Text);
        }
    }
}
