using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace ClosedXML_Tests.Excel
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class RowTests
    {

        [TestMethod]
        public void RowUsed()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 2).SetValue("Test");
            ws.Cell(1, 3).SetValue("Test");

            var fromRow = ws.Row(1).RowUsed();
            Assert.AreEqual("B1:C1", fromRow.RangeAddress.ToStringRelative());

            var fromRange = ws.Range("A1:E1").FirstRow().RowUsed();
            Assert.AreEqual("B1:C1", fromRange.RangeAddress.ToStringRelative());
        }


        [TestMethod]
        public void NoRowsUsed()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            Int32 count = 0;

            foreach (var row in ws.RowsUsed())
                count++;
            
            foreach (var row in ws.Range("A1:C3").RowsUsed())
                count++;

            Assert.AreEqual(0, count);
        }

        [TestMethod]
        public void InsertingRowsAbove1()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Rows("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Row(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

            var row1 = ws.Row(1);
            var row2 = ws.Row(2);
            var row3 = ws.Row(3);

            var rowIns = ws.Row(1).InsertRowsAbove(1).First();

            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Row(1).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Row(1).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Row(1).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, ws.Row(2).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(2).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(2).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, ws.Row(3).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green,  ws.Row(3).Cell(2).Style.Fill.BackgroundColor);
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
            Assert.AreEqual(XLColor.Green,  row2.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, row2.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Red, row3.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, row3.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, row3.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", row2.Cell(2).GetString());
        }

        [TestMethod]
        public void InsertingRowsAbove2()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Rows("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Row(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

            var row1 = ws.Row(1);
            var row2 = ws.Row(2);
            var row3 = ws.Row(3);

            var rowIns = ws.Row(2).InsertRowsAbove(1).First();

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

        [TestMethod]
        public void InsertingRowsAbove3()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Rows("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Row(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

            var row1 = ws.Row(1);
            var row2 = ws.Row(2);
            var row3 = ws.Row(3);

            var rowIns = ws.Row(3).InsertRowsAbove(1).First();

            Assert.AreEqual(XLColor.Red, ws.Row(1).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(1).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Red, ws.Row(1).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, ws.Row(2).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green,  ws.Row(2).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Yellow, ws.Row(2).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XLColor.Yellow, ws.Row(3).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Green,  ws.Row(3).Cell(2).Style.Fill.BackgroundColor);
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
    }
}
