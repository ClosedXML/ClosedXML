using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel.ConditionalFormats
{
    [TestFixture]
    public class ConditionalFormatesCompressTests
    {
        [Test]
        public void ConsecutivelyRowsCompressTest()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("B2:C2").AddConditionalFormat());
            SetFormat1(ws.Range("B4:C4").AddConditionalFormat());
            SetFormat1(ws.Range("B3:C3").AddConditionalFormat());

            ws.ConditionalFormats.Compress();

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            var format = ws.ConditionalFormats.First();
            Assert.AreEqual("B2:C4", format.Range.RangeAddress.ToStringRelative());
            Assert.AreEqual("F2", format.Values.Values.First().Value);
        }

        [Test]
        public void ConsecutivelyColumnsCompressTest()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("D2:D3").AddConditionalFormat());
            SetFormat1(ws.Range("B2:B3").AddConditionalFormat());
            SetFormat1(ws.Range("C2:C3").AddConditionalFormat());

            ws.ConditionalFormats.Compress();

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            var format = ws.ConditionalFormats.First();
            Assert.AreEqual("B2:D3", format.Range.RangeAddress.ToStringRelative());
            Assert.AreEqual("F2", format.Values.Values.First().Value);
        }

        [Test]
        public void Contains1CompressTest()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");
            
            SetFormat1(ws.Range("B11:D12").AddConditionalFormat());
            SetFormat1(ws.Range("C12:D12").AddConditionalFormat());

            ws.ConditionalFormats.Compress();

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            var format = ws.ConditionalFormats.First();
            Assert.AreEqual("B11:D12", format.Range.RangeAddress.ToStringRelative());
            Assert.AreEqual("F11", format.Values.Values.First().Value);
        }

        [Test]
        public void Contains2CompressTest()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("B14:C14").AddConditionalFormat());
            SetFormat1(ws.Range("B14:B14").AddConditionalFormat());

            ws.ConditionalFormats.Compress();

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            var format = ws.ConditionalFormats.First();
            Assert.AreEqual("B14:C14", format.Range.RangeAddress.ToStringRelative());
            Assert.AreEqual("F14", format.Values.Values.First().Value);
        }

        [Test]
        public void SuperimposedCompressTest()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("B16:D18").AddConditionalFormat());
            SetFormat1(ws.Range("B18:D19").AddConditionalFormat());

            ws.ConditionalFormats.Compress();

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            var format = ws.ConditionalFormats.First();
            Assert.AreEqual("B16:D19", format.Range.RangeAddress.ToStringRelative());
            Assert.AreEqual("F16", format.Values.Values.First().Value);
        }

        [Test]
        public void DifferentRangesNoCompressTest()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");
            
            SetFormat1(ws.Range("B7:C7").AddConditionalFormat());
            SetFormat1(ws.Range("B8:B8").AddConditionalFormat());
            SetFormat1(ws.Range("B9:C9").AddConditionalFormat());

            ws.ConditionalFormats.Compress();

            Assert.AreEqual(3, ws.ConditionalFormats.Count());
        }

        [Test]
        public void DifferentFormatNoCompressTest()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("B11:D12").AddConditionalFormat());
            SetFormat2(ws.Range("C12:D12").AddConditionalFormat());

            ws.ConditionalFormats.Compress();

            Assert.AreEqual(2, ws.ConditionalFormats.Count());
        }

        private static void SetFormat1(IXLConditionalFormat format)
        {
            format.WhenEquals("="+format.Range.FirstCell().CellRight(4).Address.ToStringRelative()).Fill.SetBackgroundColor(XLColor.Blue);
        }

        private static void SetFormat2(IXLConditionalFormat format)
        {
            format.WhenEquals(5).Fill.SetBackgroundColor(XLColor.AliceBlue);
        }
    }
}