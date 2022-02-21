using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class PageBreaksTests
    {
        [Test]
        public void RowBreaksShouldBeSorted()
        {
            var wb = new XLWorkbook();
            IXLWorksheet sheet = wb.AddWorksheet("Sheet1");

            sheet.PageSetup.AddHorizontalPageBreak(10);
            sheet.PageSetup.AddHorizontalPageBreak(12);
            sheet.PageSetup.AddHorizontalPageBreak(5);
            Assert.That(sheet.PageSetup.RowBreaks, Is.EqualTo(new[] { 5, 10, 12 }));
        }

        [Test]
        public void ColumnBreaksShouldBeSorted()
        {
            var wb = new XLWorkbook();
            IXLWorksheet sheet = wb.AddWorksheet("Sheet1");

            sheet.PageSetup.AddVerticalPageBreak(10);
            sheet.PageSetup.AddVerticalPageBreak(12);
            sheet.PageSetup.AddVerticalPageBreak(5);
            Assert.That(sheet.PageSetup.ColumnBreaks, Is.EqualTo(new[] { 5, 10, 12 }));
        }

        [Test]
        public void RowBreaksShiftWhenInsertedRowAbove()
        {
            var wb = new XLWorkbook();
            IXLWorksheet sheet = wb.AddWorksheet("Sheet1");

            sheet.PageSetup.AddHorizontalPageBreak(10);
            sheet.Row(5).InsertRowsAbove(1);
            Assert.AreEqual(11, sheet.PageSetup.RowBreaks[0]);
        }

        [Test]
        public void RowBreaksNotShiftWhenInsertedRowBelow()
        {
            var wb = new XLWorkbook();
            IXLWorksheet sheet = wb.AddWorksheet("Sheet1");

            sheet.PageSetup.AddHorizontalPageBreak(10);
            sheet.Row(15).InsertRowsAbove(1);
            Assert.AreEqual(10, sheet.PageSetup.RowBreaks[0]);
        }

        [Test]
        public void ColumnBreaksShiftWhenInsertedColumnBefore()
        {
            var wb = new XLWorkbook();
            IXLWorksheet sheet = wb.AddWorksheet("Sheet1");

            sheet.PageSetup.AddVerticalPageBreak(10);
            sheet.Column(5).InsertColumnsBefore(1);
            Assert.AreEqual(11, sheet.PageSetup.ColumnBreaks[0]);
        }

        [Test]
        public void ColumnBreaksNotShiftWhenInsertedColumnAfter()
        {
            var wb = new XLWorkbook();
            IXLWorksheet sheet = wb.AddWorksheet("Sheet1");

            sheet.PageSetup.AddVerticalPageBreak(10);
            sheet.Column(15).InsertColumnsBefore(1);
            Assert.AreEqual(10, sheet.PageSetup.ColumnBreaks[0]);
        }
    }
}
