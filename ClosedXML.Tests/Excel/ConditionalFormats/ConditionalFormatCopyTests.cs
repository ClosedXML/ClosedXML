using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML.Tests.Excel.ConditionalFormats
{
    [TestFixture]
    public class ConditionalFormatCopyTests
    {
        [Test]
        public void StylesAreCreatedDuringCopy()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");
            var format = ws.Range("A1:A1").AddConditionalFormat();
            format.WhenEquals("=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()).Fill
                  .SetBackgroundColor(XLColor.Blue);

            var wb2 = new XLWorkbook();
            var ws2 = wb2.Worksheets.Add("Sheet2");
            ws2.FirstCell().CopyFrom(ws.FirstCell());
            Assert.That(ws2.ConditionalFormats.First().Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Blue)); //Added blue style
        }

        [Test]
        public void CopyConditionalFormatSingleWorksheet()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");
            var format = ws.Range("A1:A1").AddConditionalFormat();
            format.WhenEquals("=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()).Fill
                .SetBackgroundColor(XLColor.Blue);

            ws.Cell("A1").CopyTo("B2");

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            Assert.AreEqual(2, ws.ConditionalFormats.First().Ranges.Count);
            Assert.AreEqual("A1:A1", ws.ConditionalFormats.First().Ranges.First().RangeAddress.ToString());
            Assert.AreEqual("B2:B2", ws.ConditionalFormats.First().Ranges.Last().RangeAddress.ToString());
        }

        [Test]
        public void CopyConditionalFormatSameRange()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");
            var format = ws.Range("A1:C3").AddConditionalFormat();
            format.WhenEquals("=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()).Fill
                .SetBackgroundColor(XLColor.Blue);

            ws.Cell("A1").CopyTo("B2");

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            Assert.AreEqual(1, ws.ConditionalFormats.First().Ranges.Count);
            Assert.AreEqual("A1:C3", ws.ConditionalFormats.First().Ranges.First().RangeAddress.ToString());
        }

        [Test]
        public void CopyConditionalFormatsDifferentWorksheets()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            var format = ws1.Range("A1:A1").AddConditionalFormat();
            format.WhenEquals("=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()).Fill
                .SetBackgroundColor(XLColor.Blue);
            var ws2 = wb.Worksheets.Add("Sheet2");
            var otherCell = ws2.Cell("B2");

            ws1.Cell("A1").CopyTo(otherCell);

            Assert.AreEqual(1, ws1.ConditionalFormats.Count());
            Assert.AreEqual(1, ws2.ConditionalFormats.Count());
            Assert.AreEqual(1, ws1.ConditionalFormats.First().Ranges.Count);
            Assert.AreEqual(1, ws2.ConditionalFormats.First().Ranges.Count);
            Assert.AreEqual("Sheet1", ws1.ConditionalFormats.First().Ranges.First().Worksheet.Name);
            Assert.AreEqual("Sheet2", ws2.ConditionalFormats.First().Ranges.First().Worksheet.Name);
            Assert.AreEqual("A1:A1", ws1.ConditionalFormats.First().Ranges.First().RangeAddress.ToString());
            Assert.AreEqual("B2:B2", ws2.ConditionalFormats.First().Ranges.First().RangeAddress.ToString());
        }

        [Test]
        public void FullCopyConditionalFormatSameWorksheet()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            var format = (XLConditionalFormat)ws1.Range("A1:A1").AddConditionalFormat();
            format.WhenEquals("=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()).Fill
                .SetBackgroundColor(XLColor.Blue);

            TestDelegate action = () => format.CopyTo(ws1);

            Assert.Throws(typeof(InvalidOperationException), action);
        }

        [Test]
        public void FullCopyConditionalFormatDifferentWorksheets()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            var format = (XLConditionalFormat)ws1.Range("A1:C3").AddConditionalFormat();
            format.WhenEquals("=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()).Fill
                .SetBackgroundColor(XLColor.Blue);
            var ws2 = wb.Worksheets.Add("Sheet2");

            format.CopyTo(ws2);

            Assert.AreEqual(1, ws1.ConditionalFormats.Count());
            Assert.AreEqual(1, ws2.ConditionalFormats.Count());
            Assert.AreEqual(1, ws1.ConditionalFormats.First().Ranges.Count);
            Assert.AreEqual(1, ws2.ConditionalFormats.First().Ranges.Count);
            Assert.AreEqual("Sheet1!A1:C3", ws1.ConditionalFormats.First().Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true));
            Assert.AreEqual("Sheet2!A1:C3", ws2.ConditionalFormats.First().Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true));
        }
    }
}
