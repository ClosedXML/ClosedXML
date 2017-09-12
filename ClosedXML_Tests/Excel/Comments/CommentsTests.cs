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
    }
}
