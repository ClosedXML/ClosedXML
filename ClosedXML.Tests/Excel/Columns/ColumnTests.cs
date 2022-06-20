using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.Columns
{
    [TestFixture]
    public class ColumnTests
    {
        [Test]
        public void ColumnUsed()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(2, 1).SetValue("Test");
            ws.Cell(3, 1).SetValue("Test");

            var fromColumn = ws.Column(1).ColumnUsed();
            Assert.AreEqual("A2:A3", fromColumn.RangeAddress.ToStringRelative());

            var fromRange = ws.Range("A1:A5").FirstColumn().ColumnUsed();
            Assert.AreEqual("A2:A3", fromRange.RangeAddress.ToStringRelative());
        }

        [Test]
        public void ColumnsUsedIsFast()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().SetValue("Hello world!");
            var columnsUsed = ws.Row(1).AsRange().ColumnsUsed();
            Assert.AreEqual(1, columnsUsed.Count());
        }

        [Test]
        public void CopyColumn()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Test").Style.Font.SetBold();
            ws.FirstColumn().CopyTo(ws.Column(2));

            Assert.IsTrue(ws.Cell("B1").Style.Font.Bold);
        }

        [Test]
        public void InsertingColumnsBefore1()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Columns("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Column(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

            var column1 = ws.Column(1);
            var column2 = ws.Column(2);
            var column3 = ws.Column(3);

            var columnIns = ws.Column(1).InsertColumnsBefore(1).First();

            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Column(1).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Column(1).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Column(1).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, ws.Column(2).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Column(2).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Column(2).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, ws.Column(3).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green, ws.Column(3).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, ws.Column(3).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, ws.Column(4).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Column(4).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Column(4).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", ws.Column(3).Cell(2).GetString());

            Assert.AreEqual(ws.Style.Fill.BackgroundColor, columnIns.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(ws.Style.Fill.BackgroundColor, columnIns.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(ws.Style.Fill.BackgroundColor, columnIns.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, column1.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, column1.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, column1.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, column2.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green, column2.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, column2.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, column3.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, column3.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, column3.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", column2.Cell(2).GetString());
        }

        [Test]
        public void InsertingColumnsBefore2()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Columns("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Column(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

            var column1 = ws.Column(1);
            var column2 = ws.Column(2);
            var column3 = ws.Column(3);

            var columnIns = ws.Column(2).InsertColumnsBefore(1).First();

            Assert.AreEqual(XLColor.Red, ws.Column(1).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Column(1).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Column(1).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, ws.Column(2).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Column(2).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Column(2).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, ws.Column(3).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green, ws.Column(3).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, ws.Column(3).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, ws.Column(4).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Column(4).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Column(4).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", ws.Column(3).Cell(2).GetString());

            Assert.AreEqual(XLColor.Red, columnIns.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, columnIns.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, columnIns.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, column1.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, column1.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, column1.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, column2.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green, column2.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, column2.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, column3.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, column3.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, column3.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", column2.Cell(2).GetString());
        }

        [Test]
        public void InsertingColumnsBefore3()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Columns("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Column(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

            var column1 = ws.Column(1);
            var column2 = ws.Column(2);
            var column3 = ws.Column(3);

            var columnIns = ws.Column(3).InsertColumnsBefore(1).First();

            Assert.AreEqual(XLColor.Red, ws.Column(1).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Column(1).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Column(1).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, ws.Column(2).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green, ws.Column(2).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, ws.Column(2).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, ws.Column(3).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green, ws.Column(3).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, ws.Column(3).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, ws.Column(4).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Column(4).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Column(4).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", ws.Column(2).Cell(2).GetString());

            Assert.AreEqual(XLColor.Yellow, columnIns.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green, columnIns.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, columnIns.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, column1.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, column1.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, column1.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, column2.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green, column2.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, column2.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, column3.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, column3.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, column3.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", column2.Cell(2).GetString());
        }

        [Test]
        public void NoColumnsUsed()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            var count = 0;

            foreach (var row in ws.ColumnsUsed())
            {
                count++;
            }

            foreach (var row in ws.Range("A1:C3").ColumnsUsed())
            {
                count++;
            }

            Assert.AreEqual(0, count);
        }

        [Test]
        public void UngroupFromAll()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.AddWorksheet("Sheet1");
            ws.Columns(1, 2).Group();
            ws.Columns(1, 2).Ungroup(true);
        }

        [Test]
        public void LastColumnUsed()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "A1";
            ws.Cell("B1").Value = "B1";
            ws.Cell("A2").Value = "A2";
            var lastCoUsed = ws.LastColumnUsed().ColumnNumber();
            Assert.AreEqual(2, lastCoUsed);
        }

        [Test]
        public void NegativeColumnNumberIsInvalid()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.AddWorksheet("Sheet1") as XLWorksheet;

            var column = new XLColumn(ws, -1);

            Assert.IsFalse(column.RangeAddress.IsValid);
        }
    }
}