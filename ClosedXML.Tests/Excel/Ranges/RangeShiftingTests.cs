using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Ranges
{
    public class RangeShiftingTests
    {
        [Test]
        public void CellReferenceRemainAfterColumnDeleted()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var d4 = ws.Cell("D4");

                ws.Column("C").Delete();

                Assert.AreSame(d4, ws.Cell("C4"));
            }
        }

        [Test]
        public void CellReferenceRemainAfterRowDeleted()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var d4 = ws.Cell("D4");

                ws.Row(3).Delete();

                Assert.AreSame(d4, ws.Cell("D3"));
            }
        }

        [Test]
        public void CellReferenceRemainAfterColumnInserted()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var d4 = ws.Cell("D4");

                ws.Column("C").InsertColumnsBefore(1);

                Assert.AreSame(d4, ws.Cell("E4"));
            }
        }

        [Test]
        public void CellReferenceRemainAfterRowInserted()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var d4 = ws.Cell("D4");

                ws.Row(3).InsertRowsAbove(1);

                Assert.AreSame(d4, ws.Cell("D5"));
            }
        }

        [Test]
        public void CellReferenceRemainAfterRangeDeleted()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var d4 = ws.Cell("D4");
                var f8 = ws.Cell("F8");

                ws.Range("B2:C5").Delete(XLShiftDeletedCells.ShiftCellsLeft);
                ws.Range("E5:F7").Delete(XLShiftDeletedCells.ShiftCellsUp);

                Assert.AreSame(d4, ws.Cell("B4"));
                Assert.AreSame(f8, ws.Cell("F5"));
            }
        }
    }
}
