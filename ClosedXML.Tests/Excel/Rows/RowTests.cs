using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class RowTests
    {
        [Test]
        public void RowsUsedIsFast()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().SetValue("Hello world!");
            var rowsUsed = ws.Column(1).AsRange().RowsUsed();
            Assert.AreEqual(1, rowsUsed.Count());
        }

        [Test]
        public void CopyRow()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Test").Style.Font.SetBold();
            ws.FirstRow().CopyTo(ws.Row(2));

            Assert.IsTrue(ws.Cell("A2").Style.Font.Bold);
        }

        [Test]
        public void InsertingRowsAbove1()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

            ws.Rows("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Row(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

            IXLRow row1 = ws.Row(1);
            IXLRow row2 = ws.Row(2);
            IXLRow row3 = ws.Row(3);

            IXLRow rowIns = ws.Row(1).InsertRowsAbove(1).First();

            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Row(1).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Row(1).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Row(1).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, ws.Row(2).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(2).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(2).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, ws.Row(3).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green, ws.Row(3).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, ws.Row(3).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", ws.Row(3).Cell(2).GetString());

            Assert.AreEqual(ws.Style.Fill.BackgroundColor, rowIns.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(ws.Style.Fill.BackgroundColor, rowIns.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(ws.Style.Fill.BackgroundColor, rowIns.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, row1.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, row1.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, row1.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, row2.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green, row2.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, row2.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, row3.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, row3.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, row3.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", row2.Cell(2).GetString());
        }

        [Test]
        public void InsertingRowsAbove2()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

            ws.Rows("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Row(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

            IXLRow row1 = ws.Row(1);
            IXLRow row2 = ws.Row(2);
            IXLRow row3 = ws.Row(3);

            IXLRow rowIns = ws.Row(2).InsertRowsAbove(1).First();

            Assert.AreEqual(XLColor.Red, ws.Row(1).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(1).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(1).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, ws.Row(2).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(2).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(2).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, ws.Row(3).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green, ws.Row(3).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, ws.Row(3).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", ws.Row(3).Cell(2).GetString());

            Assert.AreEqual(XLColor.Red, rowIns.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, rowIns.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, rowIns.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, row1.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, row1.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, row1.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, row2.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green, row2.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, row2.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, row3.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, row3.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, row3.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", row2.Cell(2).GetString());
        }

        [Test]
        public void InsertingRowsAbove3()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

            ws.Rows("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Row(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

            IXLRow row1 = ws.Row(1);
            IXLRow row2 = ws.Row(2);
            IXLRow row3 = ws.Row(3);

            IXLRow rowIns = ws.Row(3).InsertRowsAbove(1).First();

            Assert.AreEqual(XLColor.Red, ws.Row(1).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(1).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(1).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, ws.Row(2).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green, ws.Row(2).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, ws.Row(2).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, ws.Row(3).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green, ws.Row(3).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, ws.Row(3).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", ws.Row(2).Cell(2).GetString());

            Assert.AreEqual(XLColor.Yellow, rowIns.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green, rowIns.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, rowIns.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, row1.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, row1.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, row1.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, row2.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green, row2.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, row2.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, row3.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, row3.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, row3.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", row2.Cell(2).GetString());
        }

        [Test]
        public void InsertingRowsAbove4()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");

                ws.Row(2).Height = 15;
                ws.Row(3).Height = 20;
                ws.Row(4).Height = 25;
                ws.Row(5).Height = 35;

                ws.Row(2).FirstCell().SetValue("Row height: 15");
                ws.Row(3).FirstCell().SetValue("Row height: 20");
                ws.Row(4).FirstCell().SetValue("Row height: 25");
                ws.Row(5).FirstCell().SetValue("Row height: 35");

                ws.Range("3:3").InsertRowsAbove(1);

                Assert.AreEqual(15, ws.Row(2).Height);
                Assert.AreEqual(20, ws.Row(4).Height);
                Assert.AreEqual(25, ws.Row(5).Height);
                Assert.AreEqual(35, ws.Row(6).Height);

                Assert.AreEqual(20, ws.Row(3).Height);
                ws.Row(3).ClearHeight();
                Assert.AreEqual(ws.RowHeight, ws.Row(3).Height);
            }
        }

        [Test]
        public void NoRowsUsed()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            Int32 count = 0;

            foreach (IXLRow row in ws.RowsUsed())
                count++;

            foreach (IXLRangeRow row in ws.Range("A1:C3").RowsUsed())
                count++;

            Assert.AreEqual(0, count);
        }

        [Test]
        public void RowUsed()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 2).SetValue("Test");
            ws.Cell(1, 3).SetValue("Test");

            IXLRangeRow fromRow = ws.Row(1).RowUsed();
            Assert.AreEqual("B1:C1", fromRow.RangeAddress.ToStringRelative());

            IXLRangeRow fromRange = ws.Range("A1:E1").FirstRow().RowUsed();
            Assert.AreEqual("B1:C1", fromRange.RangeAddress.ToStringRelative());
        }

        [Test]
        public void RowsUsedWithDataValidation()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().SetValue("Hello world!");
            ws.Range("A1:A100").CreateDataValidation().WholeNumber.EqualTo(1);

            var range = ws.Column(1).AsRange();

            Assert.AreEqual(100, range.RowsUsed(XLCellsUsedOptions.DataValidation).Count());
            Assert.AreEqual(100, range.RowsUsed(XLCellsUsedOptions.All).Count());
        }

        [Test]
        public void RowsUsedWithConditionalFormatting()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().SetValue("Hello world!");
            ws.Range("A1:A100").AddConditionalFormat().WhenStartsWith("Hell").Fill.SetBackgroundColor(XLColor.Red).Font.SetFontColor(XLColor.White);

            var range = ws.Column(1).AsRange();

            Assert.AreEqual(100, range.RowsUsed(XLCellsUsedOptions.ConditionalFormats).Count());
            Assert.AreEqual(100, range.RowsUsed(XLCellsUsedOptions.All).Count());
        }

        [Test]
        public void UngroupFromAll()
        {
            IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");
            ws.Rows(1, 2).Group();
            ws.Rows(1, 2).Ungroup(true);
        }

        [Test]
        public void NegativeRowNumberIsInvalid()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1") as XLWorksheet;

            var row = new XLRow(ws, -1);

            Assert.IsFalse(row.RangeAddress.IsValid);
        }

        [Test]
        public void DeleteRowOnWorksheetWithComment()
        {
            var ws = new XLWorkbook().AddWorksheet();
            ws.Cell(4, 1).GetComment().AddText("test");
            ws.Column(1).Width = 100;
            Assert.DoesNotThrow(() => ws.Row(1).Delete());
        }
    }
}
