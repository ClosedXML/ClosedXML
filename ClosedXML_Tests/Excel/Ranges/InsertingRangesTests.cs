using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class InsertingRangesTests
    {
        [Test]
        public void InsertingColumnsPreservesFormatting()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");
            IXLColumn column1 = ws.Column(1);
            column1.Style.Fill.SetBackgroundColor(XLColor.FrenchLilac);
            column1.Cell(2).Style.Fill.SetBackgroundColor(XLColor.Fulvous);
            IXLColumn column2 = ws.Column(2);
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

        [Test]
        public void InsertingRowsAbove()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            ws.Cell("B3").SetValue("X")
                .CellBelow().SetValue("B");

            IXLRangeRow r = ws.Range("B4").InsertRowsAbove(1).First();
            r.Cell(1).SetValue("A");

            Assert.AreEqual("X", ws.Cell("B3").GetString());
            Assert.AreEqual("A", ws.Cell("B4").GetString());
            Assert.AreEqual("B", ws.Cell("B5").GetString());
        }

        [Test]
        public void InsertingRowsPreservesFormatting()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");
            IXLRow row1 = ws.Row(1);
            row1.Style.Fill.SetBackgroundColor(XLColor.FrenchLilac);
            row1.Cell(2).Style.Fill.SetBackgroundColor(XLColor.Fulvous);
            IXLRow row2 = ws.Row(2);
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

        [Test]
        public void InsertingRowsPreservesComments()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Cell("A1").SetValue("Insert Below");
            ws.Cell("A2").SetValue("Already existing cell");
            ws.Cell("A3").SetValue("Cell with comment").Comment.AddText("Comment here");

            ws.Row(1).InsertRowsBelow(2);
            Assert.AreEqual("Comment here", ws.Cell("A5").Comment.Text);
        }

        [Test]
        public void InsertingColumnsPreservesComments()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Cell("A1").SetValue("Insert to the right");
            ws.Cell("B1").SetValue("Already existing cell");
            ws.Cell("C1").SetValue("Cell with comment").Comment.AddText("Comment here");

            ws.Column(1).InsertColumnsAfter(2);
            Assert.AreEqual("Comment here", ws.Cell("E1").Comment.Text);
        }
    }
}
