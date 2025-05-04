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
            Assert.That(rowsUsed.Count(), Is.EqualTo(1));
        }

        [Test]
        public void CopyRow()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Test").Style.Font.SetBold();
            ws.FirstRow().CopyTo(ws.Row(2));

            Assert.That(ws.Cell("A2").Style.Font.Bold, Is.True);
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

            Assert.That(ws.Row(1).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));
            Assert.That(ws.Row(1).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));
            Assert.That(ws.Row(1).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));

            Assert.Multiple(() =>
            {
                Assert.That(ws.Row(2).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Row(2).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Row(2).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(ws.Row(3).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
                Assert.That(ws.Row(3).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(ws.Row(3).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

                Assert.That(ws.Row(4).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Row(4).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Row(4).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(ws.Row(3).Cell(2).GetText(), Is.EqualTo("X"));

                Assert.That(rowIns.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));
            });
            Assert.That(rowIns.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));
            Assert.That(rowIns.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));

            Assert.That(row1.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.Multiple(() =>
            {
                Assert.That(row1.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(row1.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(row2.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
                Assert.That(row2.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(row2.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

                Assert.That(row3.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(row3.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(row3.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(row2.Cell(2).GetText(), Is.EqualTo("X"));
            });
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

            Assert.Multiple(() =>
            {
                Assert.That(ws.Row(1).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Row(1).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Row(1).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(ws.Row(2).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Row(2).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Row(2).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(ws.Row(3).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
                Assert.That(ws.Row(3).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(ws.Row(3).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

                Assert.That(ws.Row(4).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Row(4).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Row(4).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(ws.Row(3).Cell(2).GetText(), Is.EqualTo("X"));

                Assert.That(rowIns.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(rowIns.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(rowIns.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(row1.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(row1.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(row1.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(row2.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
                Assert.That(row2.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(row2.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

                Assert.That(row3.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(row3.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(row3.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(row2.Cell(2).GetText(), Is.EqualTo("X"));
            });
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

            Assert.Multiple(() =>
            {
                Assert.That(ws.Row(1).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Row(1).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Row(1).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(ws.Row(2).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
                Assert.That(ws.Row(2).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(ws.Row(2).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

                Assert.That(ws.Row(3).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
                Assert.That(ws.Row(3).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(ws.Row(3).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

                Assert.That(ws.Row(4).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Row(4).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Row(4).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(ws.Row(2).Cell(2).GetText(), Is.EqualTo("X"));

                Assert.That(rowIns.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
                Assert.That(rowIns.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(rowIns.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

                Assert.That(row1.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(row1.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(row1.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(row2.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
                Assert.That(row2.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(row2.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

                Assert.That(row3.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(row3.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(row3.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(row2.Cell(2).GetText(), Is.EqualTo("X"));
            });
        }

        [Test]
        public void InsertingRowsAbove4()
        {
            using var wb = new XLWorkbook();
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

            Assert.Multiple(() =>
            {
                Assert.That(ws.Row(2).Height, Is.EqualTo(15));
                Assert.That(ws.Row(4).Height, Is.EqualTo(20));
                Assert.That(ws.Row(5).Height, Is.EqualTo(25));
                Assert.That(ws.Row(6).Height, Is.EqualTo(35));

                Assert.That(ws.Row(3).Height, Is.EqualTo(20));
            });
            ws.Row(3).ClearHeight();
            Assert.That(ws.Row(3).Height, Is.EqualTo(ws.RowHeight));
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

            Assert.That(count, Is.EqualTo(0));
        }

        [Test]
        public void RowUsed()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 2).SetValue("Test");
            ws.Cell(1, 3).SetValue("Test");

            IXLRangeRow fromRow = ws.Row(1).RowUsed();
            Assert.That(fromRow.RangeAddress.ToStringRelative(), Is.EqualTo("B1:C1"));

            IXLRangeRow fromRange = ws.Range("A1:E1").FirstRow().RowUsed();
            Assert.That(fromRange.RangeAddress.ToStringRelative(), Is.EqualTo("B1:C1"));
        }

        [Test]
        public void RowsUsedWithDataValidation()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().SetValue("Hello world!");
            ws.Range("A1:A100").CreateDataValidation().WholeNumber.EqualTo(1);

            var range = ws.Column(1).AsRange();

            Assert.Multiple(() =>
            {
                Assert.That(range.RowsUsed(XLCellsUsedOptions.DataValidation).Count(), Is.EqualTo(100));
                Assert.That(range.RowsUsed(XLCellsUsedOptions.All).Count(), Is.EqualTo(100));
            });
        }

        [Test]
        public void RowsUsedWithConditionalFormatting()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().SetValue("Hello world!");
            ws.Range("A1:A100").AddConditionalFormat().WhenStartsWith("Hell").Fill.SetBackgroundColor(XLColor.Red).Font.SetFontColor(XLColor.White);

            var range = ws.Column(1).AsRange();

            Assert.Multiple(() =>
            {
                Assert.That(range.RowsUsed(XLCellsUsedOptions.ConditionalFormats).Count(), Is.EqualTo(100));
                Assert.That(range.RowsUsed(XLCellsUsedOptions.All).Count(), Is.EqualTo(100));
            });
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

            Assert.That(row.RangeAddress.IsValid, Is.False);
        }

        [Test]
        public void DeleteRowOnWorksheetWithComment()
        {
            var ws = new XLWorkbook().AddWorksheet();
            ws.Cell(4, 1).GetComment().AddText("test");
            ws.Column(1).Width = 100;
            Assert.DoesNotThrow(() => ws.Row(1).Delete());
        }

        [Test]
        public void AssignWorksheetRowHeightWhenAllRowsChanged()
        {
            var ws = new XLWorkbook().AddWorksheet();
            var rows = ws.Rows();

            rows.Height = 30;

            Assert.Multiple(() =>
            {
                Assert.That(ws.Row(11).Height, Is.EqualTo(30).Within(XLHelper.Epsilon));
                Assert.That(ws.RowHeight, Is.EqualTo(30).Within(XLHelper.Epsilon));
            });
        }

        [Test]
        public void PreserveWorksheetRowHeightWhenNotAllRowsChanged()
        {
            var ws = new XLWorkbook().AddWorksheet();
            var defaultRowHeight = ws.RowHeight;
            var rows = ws.Rows(1, XLHelper.MaxRowNumber);

            rows.Height = 30;

            Assert.Multiple(() =>
            {
                Assert.That(ws.Row(11).Height, Is.EqualTo(30).Within(XLHelper.Epsilon));
                Assert.That(ws.RowHeight, Is.EqualTo(defaultRowHeight).Within(XLHelper.Epsilon));
            });
        }

        [Test]
        public void PreserveWorksheetRowHeightWhenUsedRowsChanged()
        {
            var ws = new XLWorkbook().AddWorksheet();
            ws.Cells("A1:E5").Value = "Not empty";
            var defaultRowHeight = ws.RowHeight;
            var rows = ws.RowsUsed(XLCellsUsedOptions.Contents);

            rows.Height = 30;

            Assert.Multiple(() =>
            {
                Assert.That(ws.Row(3).Height, Is.EqualTo(30).Within(XLHelper.Epsilon));
                Assert.That(ws.Row(11).Height, Is.EqualTo(defaultRowHeight).Within(XLHelper.Epsilon));
                Assert.That(ws.RowHeight, Is.EqualTo(defaultRowHeight).Within(XLHelper.Epsilon));
            });
        }
    }
}
