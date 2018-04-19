using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML_Tests.Excel.Ranges
{
    [TestFixture]
    public class RangesConsolidationTests
    {
        [Test]
        public void ConsolidateRangesSameWorksheet()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            var ranges = new XLRanges();
            ranges.Add(ws.Range("A1:E3"));
            ranges.Add(ws.Range("A4:B10"));
            ranges.Add(ws.Range("E2:F12"));
            ranges.Add(ws.Range("C6:I8"));
            ranges.Add(ws.Range("G9:G9"));
            ranges.Add(ws.Range("C9:D9"));
            ranges.Add(ws.Range("H9:H9"));
            ranges.Add(ws.Range("I9:I13"));
            ranges.Add(ws.Range("C4:D5"));

            var consRanges = ranges.Consolidate().ToList();

            Assert.AreEqual(6, consRanges.Count);
            Assert.AreEqual("A1:E9",   consRanges[0].RangeAddress.ToString());
            Assert.AreEqual("F2:F12",  consRanges[1].RangeAddress.ToString());
            Assert.AreEqual("G6:I9",   consRanges[2].RangeAddress.ToString());
            Assert.AreEqual("A10:B10", consRanges[3].RangeAddress.ToString());
            Assert.AreEqual("E10:E12", consRanges[4].RangeAddress.ToString());
            Assert.AreEqual("I10:I13", consRanges[5].RangeAddress.ToString());
        }

        [Test]
        public void ConsolidateWideRangesSameWorksheet()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            var ranges = new XLRanges();
            ranges.Add(ws.Row(5));
            ranges.Add(ws.Row(7));
            ranges.Add(ws.Row(6));
            ranges.Add(ws.Column("D"));
            ranges.Add(ws.Column("F"));
            ranges.Add(ws.Column("E"));

            var consRanges = ranges.Consolidate().ToList();

            Assert.AreEqual(3, consRanges.Count);
            Assert.AreEqual("D1:F1048576", consRanges[0].RangeAddress.ToString());
            Assert.AreEqual("A5:C7", consRanges[1].RangeAddress.ToString());
            Assert.AreEqual("G5:XFD7", consRanges[2].RangeAddress.ToString());
        }

        [Test]
        public void ConsolidateRangesDifferentWorksheets()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            var ws2 = wb.Worksheets.Add("Sheet2");
            var ranges = new XLRanges();
            ranges.Add(ws1.Range("A1:E3"));
            ranges.Add(ws1.Range("A4:B10"));
            ranges.Add(ws1.Range("E2:F12"));
            ranges.Add(ws1.Range("C6:I8"));
            ranges.Add(ws1.Range("G9:G9"));

            ranges.Add(ws2.Row(5));
            ranges.Add(ws2.Row(7));
            ranges.Add(ws2.Row(6));
            ranges.Add(ws2.Column("D"));
            ranges.Add(ws2.Column("F"));
            ranges.Add(ws2.Column("E"));

            ranges.Add(ws1.Range("C9:D9"));
            ranges.Add(ws1.Range("H9:H9"));
            ranges.Add(ws1.Range("I9:I13"));
            ranges.Add(ws1.Range("C4:D5"));
            
            var consRanges = ranges.Consolidate().ToList();

            Assert.AreEqual(9, consRanges.Count);
            Assert.AreEqual("Sheet1!$A$1:$E$9", consRanges[0].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true));
            Assert.AreEqual("Sheet1!$F$2:$F$12", consRanges[1].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true));
            Assert.AreEqual("Sheet1!$G$6:$I$9", consRanges[2].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true));
            Assert.AreEqual("Sheet1!$A$10:$B$10", consRanges[3].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true));
            Assert.AreEqual("Sheet1!$E$10:$E$12", consRanges[4].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true));
            Assert.AreEqual("Sheet1!$I$10:$I$13", consRanges[5].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true));

            Assert.AreEqual("Sheet2!$D$1:$F$1048576", consRanges[6].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true));
            Assert.AreEqual("Sheet2!$A$5:$C$7", consRanges[7].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true));
            Assert.AreEqual("Sheet2!$G$5:$XFD$7", consRanges[8].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true));
        }
    }
}
