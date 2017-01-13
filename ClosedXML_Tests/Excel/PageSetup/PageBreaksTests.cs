using System.Diagnostics;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel
{
    [TestFixture]
    public class PageBreaksTests
    {
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