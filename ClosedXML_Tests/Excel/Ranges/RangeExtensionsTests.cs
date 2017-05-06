using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel.Ranges
{
    [TestFixture]
    public class RangeExtensionsTests
    {
        [Test]
        public void RelativeTest()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");
            var range = ws.Range("C11:D12");
            var baseRange = ws.Range("B10:E13");
            var targetRange = ws.Range("C3:F6");

            var assert = range.Relative(baseRange, targetRange);
            Assert.AreEqual("D4", assert.RangeAddress.FirstAddress.ToStringRelative());
            Assert.AreEqual("E5", assert.RangeAddress.LastAddress.ToStringRelative());
        }        

        [Test]
        public void CropTest()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");
            var range = ws.Range("B3:D8");
            var crop = ws.Range("A5:E6");
            var assert = range.Crop(crop);

            Assert.AreEqual("B5", assert.RangeAddress.FirstAddress.ToStringRelative());
            Assert.AreEqual("D6", assert.RangeAddress.LastAddress.ToStringRelative());
        }        
    }
}