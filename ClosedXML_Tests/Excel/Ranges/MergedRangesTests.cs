using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class MergedRangesTests
    {
        [Test]
        public void LastCellFromMerge()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");
            ws.Range("B2:D4").Merge();

            string first = ws.FirstCellUsed(true).Address.ToStringRelative();
            string last = ws.LastCellUsed(true).Address.ToStringRelative();

            Assert.AreEqual("B2", first);
            Assert.AreEqual("D4", last);
        }
    }
}