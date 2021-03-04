using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Misc
{
    [TestFixture]
    public class HyperlinkTests
    {
        [Test]
        public void TestHyperlinks()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.Worksheets.Add("Sheet1");
                var ws2 = wb.Worksheets.Add("Sheet2");

                var targetCell = ws2.Cell("A1");
                var targetRange = ws2.Range("A1", "B1");

                var linkCell1 = ws1.Cell("A1");
                linkCell1.Value = "Link to IXLCell";
                linkCell1.SetHyperlink(new XLHyperlink(targetCell));
                Assert.AreEqual("Sheet2!A1", linkCell1.GetHyperlink().InternalAddress);

                var linkRange1 = ws1.Cell("A2");
                linkRange1.Value = "Link to IXLRangeBase";
                linkRange1.SetHyperlink(new XLHyperlink(targetRange));
                Assert.AreEqual("Sheet2!A1:B1", linkRange1.GetHyperlink().InternalAddress);
            }
        }
    }
}
