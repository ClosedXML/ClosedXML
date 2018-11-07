using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML_Tests.Excel.Comments
{
    public class CommentsTests
    {
        [Test]
        public void CanGetColorFromIndex81()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\CommentsWithIndexedColor81.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                var c = ws.FirstCellUsed();

                var xlColor = c.Comment.Style.ColorsAndLines.LineColor;
                Assert.AreEqual(XLColorType.Indexed, xlColor.ColorType);
                Assert.AreEqual(81, xlColor.Indexed);

                var color = xlColor.Color.ToHex();
                Assert.AreEqual("FF000000", color);
            }
        }

        [Test]
        public void AddingCommentDoesNotAffectCollections()
        {
            var ws = new XLWorkbook().AddWorksheet() as XLWorksheet;
            ws.Cell("A1").SetValue(10);
            ws.Cell("A4").SetValue(10);
            ws.Cell("A5").SetValue(10);

            ws.Rows("1,4").Height = 20;

            Assert.AreEqual(2, ws.Internals.RowsCollection.Count);
            Assert.AreEqual(3, ws.Internals.CellsCollection.RowsCollection.SelectMany(r => r.Value.Values).Count());

            ws.Cell("A4").Comment.AddText("Comment");
            Assert.AreEqual(2, ws.Internals.RowsCollection.Count);
            Assert.AreEqual(3, ws.Internals.CellsCollection.RowsCollection.SelectMany(r => r.Value.Values).Count());

            ws.Row(1).Delete();
            Assert.AreEqual(1, ws.Internals.RowsCollection.Count);
            Assert.AreEqual(2, ws.Internals.CellsCollection.RowsCollection.SelectMany(r => r.Value.Values).Count());
        }

        [Test]
        public void CopyCommentStyle()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                string strExcelComment = "1) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
                strExcelComment = strExcelComment + "1) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
                strExcelComment = strExcelComment + "2) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
                strExcelComment = strExcelComment + "3) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
                strExcelComment = strExcelComment + "4) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
                strExcelComment = strExcelComment + "5) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
                strExcelComment = strExcelComment + "6) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
                strExcelComment = strExcelComment + "7) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
                strExcelComment = strExcelComment + "8) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
                strExcelComment = strExcelComment + "9) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;

                var cell = ws.Cell(2, 2).SetValue("Comment 1");

                cell.Comment
                    .SetVisible(false)
                    .AddText(strExcelComment);

                cell.Comment
                    .Style
                    .Alignment
                    .SetAutomaticSize();

                cell.Comment
                    .Style
                    .ColorsAndLines
                    .SetFillColor(XLColor.Red);

                ws.Row(1).InsertRowsAbove(1);

                Action<IXLCell> validate = c =>
                {
                    Assert.IsTrue(c.Comment.Style.Alignment.AutomaticSize);
                    Assert.AreEqual(XLColor.Red, c.Comment.Style.ColorsAndLines.FillColor);
                };

                validate(ws.Cell("B3"));

                ws.Column(1).InsertColumnsBefore(2);

                validate(ws.Cell("D3"));

                ws.Column(1).Delete();

                validate(ws.Cell("C3"));

                ws.Row(1).Delete();

                validate(ws.Cell("C2"));
            }
        }
    }
}
