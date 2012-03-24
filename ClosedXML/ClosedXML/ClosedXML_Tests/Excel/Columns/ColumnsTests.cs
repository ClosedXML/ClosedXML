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
    public class ColumnTests
    {

        [TestMethod]
        public void ColumnUsed()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(2, 1).SetValue("Test");
            ws.Cell(3, 1).SetValue("Test");

            var fromColumn = ws.Column(1).ColumnUsed();
            Assert.AreEqual("A2:A3", fromColumn.RangeAddress.ToStringRelative());

            var fromRange = ws.Range("A1:A5").FirstColumn().ColumnUsed();
            Assert.AreEqual("A2:A3", fromRange.RangeAddress.ToStringRelative());
        }

        [TestMethod]
        public void NoColumnsUsed()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            Int32 count = 0;

            foreach (var row in ws.ColumnsUsed())
                count++;

            foreach (var row in ws.Range("A1:C3").ColumnsUsed())
                count++;

            Assert.AreEqual(0, count);
        }

        [TestMethod]
        public void InsertingColumnsBefore1()
        {
            var wb = new XLWorkbook();
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

        [TestMethod]
        public void InsertingColumnsBefore2()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Columns("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Column(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

            var column1 = ws.Column(1);
            var column2 = ws.Column(2);
            var column3 = ws.Column(3);

            var columnIns = ws.Column(2).InsertColumnsBefore(1).First();
            wb.SaveAs(@"D:\Excel Files\ForTesting\Sandbox.xlsx");

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

        [TestMethod]
        public void InsertingColumnsBefore3()
        {
            var wb = new XLWorkbook();
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
    }
}
