using System;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel
{
    /// <summary>
    ///     Summary description for UnitTest1
    /// </summary>
    [TestFixture]
    public class RowTests
    {
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
        public void UngroupFromAll()
        {
            IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");
            ws.Rows(1, 2).Group();
            ws.Rows(1, 2).Ungroup(true);
        }
    }
}