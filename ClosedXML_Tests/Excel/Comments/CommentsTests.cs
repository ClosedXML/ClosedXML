using ClosedXML.Excel;
using NUnit.Framework;
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
    }
}
