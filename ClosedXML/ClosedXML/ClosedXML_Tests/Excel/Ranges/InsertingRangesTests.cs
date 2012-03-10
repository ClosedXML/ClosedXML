using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System;
using System.IO;

namespace ClosedXML_Tests
{
    [TestClass()]
    public class InsertingRangesTests
    {
        [TestMethod]
        public void InsertingRowsPreservesFormatting()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");
            var row1 = ws.Row(1);
            row1.Style.Fill.SetBackgroundColor(XLColor.FrenchLilac);
            row1.Cell(2).Style.Fill.SetBackgroundColor(XLColor.Fulvous);
            var row2 = ws.Row(2);
            row2.Style.Fill.SetBackgroundColor(XLColor.Xanadu);
            row2.Cell(2).Style.Fill.SetBackgroundColor(XLColor.MacaroniAndCheese);

            row1.InsertRowsBelow(1);
            row1.InsertRowsAbove(1);
            row2.InsertRowsAbove(1);

            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Row(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FrenchLilac, ws.Row(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FrenchLilac, ws.Row(3).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FrenchLilac, ws.Row(4).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Xanadu, ws.Row(5).Style.Fill.BackgroundColor);

            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Cell(1, 2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Fulvous, ws.Cell(2, 2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Fulvous, ws.Cell(3, 2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Fulvous, ws.Cell(4, 2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.MacaroniAndCheese, ws.Cell(5, 2).Style.Fill.BackgroundColor);
        }

        [TestMethod]
        public void InsertingColumnsPreservesFormatting()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");
            var column1 = ws.Column(1);
            column1.Style.Fill.SetBackgroundColor(XLColor.FrenchLilac);
            column1.Cell(2).Style.Fill.SetBackgroundColor(XLColor.Fulvous);
            var column2 = ws.Column(2);
            column2.Style.Fill.SetBackgroundColor(XLColor.Xanadu);
            column2.Cell(2).Style.Fill.SetBackgroundColor(XLColor.MacaroniAndCheese);

            column1.InsertColumnsAfter(1);
            column1.InsertColumnsBefore(1);
            column2.InsertColumnsBefore(1);

            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Column(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FrenchLilac, ws.Column(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FrenchLilac, ws.Column(3).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FrenchLilac, ws.Column(4).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Xanadu, ws.Column(5).Style.Fill.BackgroundColor);

            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Cell(2, 1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Fulvous, ws.Cell(2, 2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Fulvous, ws.Cell(2, 3).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Fulvous, ws.Cell(2, 4).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.MacaroniAndCheese, ws.Cell(2, 5).Style.Fill.BackgroundColor);
        }

        [TestMethod]
        public void InsertingRowsAbove()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");

            ws.Cell("B3").SetValue("X")
                .CellBelow().SetValue("B");

            var r = ws.Range("B4").InsertRowsAbove(1).First();
            r.Cell(1).SetValue("A");

            Assert.AreEqual("X", ws.Cell("B3").GetString());
            Assert.AreEqual("A", ws.Cell("B4").GetString());
            Assert.AreEqual("B", ws.Cell("B5").GetString());
        }
    }
}
