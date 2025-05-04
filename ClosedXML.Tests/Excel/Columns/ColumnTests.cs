using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class ColumnTests
    {
        [Test]
        public void ColumnUsed()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(2, 1).SetValue("Test");
            ws.Cell(3, 1).SetValue("Test");

            IXLRangeColumn fromColumn = ws.Column(1).ColumnUsed();
            Assert.That(fromColumn.RangeAddress.ToStringRelative(), Is.EqualTo("A2:A3"));

            IXLRangeColumn fromRange = ws.Range("A1:A5").FirstColumn().ColumnUsed();
            Assert.That(fromRange.RangeAddress.ToStringRelative(), Is.EqualTo("A2:A3"));
        }

        [Test]
        public void ColumnsUsedIsFast()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().SetValue("Hello world!");
            var columnsUsed = ws.Row(1).AsRange().ColumnsUsed();
            Assert.That(columnsUsed.Count(), Is.EqualTo(1));
        }

        [Test]
        public void CopyColumn()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Test").Style.Font.SetBold();
            ws.FirstColumn().CopyTo(ws.Column(2));

            Assert.That(ws.Cell("B1").Style.Font.Bold, Is.True);
        }

        [Test]
        public void InsertingColumnsBefore1()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

            ws.Columns("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Column(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

            IXLColumn column1 = ws.Column(1);
            IXLColumn column2 = ws.Column(2);
            IXLColumn column3 = ws.Column(3);

            IXLColumn columnIns = ws.Column(1).InsertColumnsBefore(1).First();

            Assert.That(ws.Column(1).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));
            Assert.That(ws.Column(1).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));
            Assert.That(ws.Column(1).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));

            Assert.Multiple(() =>
            {
                Assert.That(ws.Column(2).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Column(2).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Column(2).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(ws.Column(3).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
                Assert.That(ws.Column(3).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(ws.Column(3).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

                Assert.That(ws.Column(4).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Column(4).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Column(4).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(ws.Column(3).Cell(2).GetText(), Is.EqualTo("X"));

                Assert.That(columnIns.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));
            });
            Assert.That(columnIns.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));
            Assert.That(columnIns.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));

            Assert.That(column1.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.Multiple(() =>
            {
                Assert.That(column1.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(column1.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(column2.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
                Assert.That(column2.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(column2.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

                Assert.That(column3.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(column3.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(column3.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(column2.Cell(2).GetText(), Is.EqualTo("X"));
            });
        }

        [Test]
        public void InsertingColumnsBefore2()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

            ws.Columns("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Column(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

            IXLColumn column1 = ws.Column(1);
            IXLColumn column2 = ws.Column(2);
            IXLColumn column3 = ws.Column(3);

            IXLColumn columnIns = ws.Column(2).InsertColumnsBefore(1).First();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Column(1).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Column(1).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Column(1).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(ws.Column(2).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Column(2).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Column(2).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(ws.Column(3).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
                Assert.That(ws.Column(3).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(ws.Column(3).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

                Assert.That(ws.Column(4).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Column(4).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Column(4).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(ws.Column(3).Cell(2).GetText(), Is.EqualTo("X"));

                Assert.That(columnIns.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(columnIns.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(columnIns.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(column1.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(column1.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(column1.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(column2.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
                Assert.That(column2.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(column2.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

                Assert.That(column3.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(column3.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(column3.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(column2.Cell(2).GetText(), Is.EqualTo("X"));
            });
        }

        [Test]
        public void InsertingColumnsBefore3()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

            ws.Columns("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Column(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

            IXLColumn column1 = ws.Column(1);
            IXLColumn column2 = ws.Column(2);
            IXLColumn column3 = ws.Column(3);

            IXLColumn columnIns = ws.Column(3).InsertColumnsBefore(1).First();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Column(1).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Column(1).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Column(1).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(ws.Column(2).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
                Assert.That(ws.Column(2).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(ws.Column(2).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

                Assert.That(ws.Column(3).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
                Assert.That(ws.Column(3).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(ws.Column(3).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

                Assert.That(ws.Column(4).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Column(4).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Column(4).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(ws.Column(2).Cell(2).GetText(), Is.EqualTo("X"));

                Assert.That(columnIns.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
                Assert.That(columnIns.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(columnIns.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

                Assert.That(column1.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(column1.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(column1.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(column2.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
                Assert.That(column2.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(column2.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

                Assert.That(column3.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(column3.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(column3.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

                Assert.That(column2.Cell(2).GetText(), Is.EqualTo("X"));
            });
        }

        [Test]
        public void NoColumnsUsed()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            Int32 count = 0;

            foreach (IXLColumn row in ws.ColumnsUsed())
                count++;

            foreach (IXLRangeColumn row in ws.Range("A1:C3").ColumnsUsed())
                count++;

            Assert.That(count, Is.EqualTo(0));
        }

        [Test]
        public void UngroupFromAll()
        {
            IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");
            ws.Columns(1, 2).Group();
            ws.Columns(1, 2).Ungroup(true);
        }

        [Test]
        public void LastColumnUsed()
        {
            IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "A1";
            ws.Cell("B1").Value = "B1";
            ws.Cell("A2").Value = "A2";
            var lastCoUsed = ws.LastColumnUsed().ColumnNumber();
            Assert.That(lastCoUsed, Is.EqualTo(2));
        }

        [Test]
        public void NegativeColumnNumberIsInvalid()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1") as XLWorksheet;

            var column = new XLColumn(ws, -1);

            Assert.That(column.RangeAddress.IsValid, Is.False);
        }

        [Test]
        public void AssignWorksheetColumnWidthWhenAllColumnsChanged()
        {
            var ws = new XLWorkbook().AddWorksheet();
            var columns = ws.Columns();

            columns.Width = 100;

            Assert.Multiple(() =>
            {
                Assert.That(ws.Column("G").Width, Is.EqualTo(100).Within(XLHelper.Epsilon));
                Assert.That(ws.ColumnWidth, Is.EqualTo(100).Within(XLHelper.Epsilon));
            });
        }

        [Test]
        public void PreserveWorksheetColumnWidthWhenNotAllColumnsChanged()
        {
            var ws = new XLWorkbook().AddWorksheet();
            var defaultColumnWidth = ws.ColumnWidth;
            var columns = ws.Columns(1, XLHelper.MaxColumnNumber);

            columns.Width = 100;

            Assert.Multiple(() =>
            {
                Assert.That(ws.Column("G").Width, Is.EqualTo(100).Within(XLHelper.Epsilon));
                Assert.That(ws.ColumnWidth, Is.EqualTo(defaultColumnWidth).Within(XLHelper.Epsilon));
            });
        }

        [Test]
        public void PreserveWorksheetColumnWidthWhenUsedColumnsChanged()
        {
            var ws = new XLWorkbook().AddWorksheet();
            ws.Cells("A1:E5").Value = "Not empty";
            var defaultColumnWidth = ws.ColumnWidth;
            var columns = ws.ColumnsUsed(XLCellsUsedOptions.Contents);

            columns.Width = 100;

            Assert.Multiple(() =>
            {
                Assert.That(ws.Column("C").Width, Is.EqualTo(100).Within(XLHelper.Epsilon));
                Assert.That(ws.Column("G").Width, Is.EqualTo(defaultColumnWidth).Within(XLHelper.Epsilon));
                Assert.That(ws.ColumnWidth, Is.EqualTo(defaultColumnWidth).Within(XLHelper.Epsilon));
            });
        }
    }
}
