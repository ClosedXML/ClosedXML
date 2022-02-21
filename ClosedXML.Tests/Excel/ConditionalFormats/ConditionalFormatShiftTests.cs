using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.ConditionalFormats
{
    [TestFixture]
    public class ConditionalFormatShiftTests
    {
        [Test]
        public void CFShiftedOnColumnInsert()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("CFShift");
                ws.Range("A1:A1").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AirForceBlue);
                ws.Range("A2:B2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AliceBlue);
                ws.Range("A3:C3").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Alizarin);
                ws.Range("B4:B6").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Almond);
                ws.Range("C7:D7").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Amaranth);
                ws.Cells("A1:D7").Value = 1;

                ws.Column(2).InsertColumnsAfter(2);
                var cf = ws.ConditionalFormats.ToArray();

                Assert.AreEqual(5, cf.Length);
                Assert.AreEqual("A1:A1", cf[0].Range.RangeAddress.ToString());
                Assert.AreEqual("A2:D2", cf[1].Range.RangeAddress.ToString());
                Assert.AreEqual("A3:E3", cf[2].Range.RangeAddress.ToString());
                Assert.AreEqual("B4:D6", cf[3].Range.RangeAddress.ToString());
                Assert.AreEqual("E7:F7", cf[4].Range.RangeAddress.ToString());
            }
        }

        [Test]
        public void CFShiftedOnRowInsert()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("CFShift");
                ws.Range("A1:A1").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AirForceBlue);
                ws.Range("B1:B2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AliceBlue);
                ws.Range("C1:C3").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Alizarin);
                ws.Range("D2:F2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Almond);
                ws.Range("G4:G5").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Amaranth);
                ws.Cells("A1:G5").Value = 1;

                ws.Row(2).InsertRowsBelow(2);
                var cf = ws.ConditionalFormats.ToArray();

                Assert.AreEqual(5, cf.Length);
                Assert.AreEqual("A1:A1", cf[0].Range.RangeAddress.ToString());
                Assert.AreEqual("B1:B4", cf[1].Range.RangeAddress.ToString());
                Assert.AreEqual("C1:C5", cf[2].Range.RangeAddress.ToString());
                Assert.AreEqual("D2:F4", cf[3].Range.RangeAddress.ToString());
                Assert.AreEqual("G6:G7", cf[4].Range.RangeAddress.ToString());
            }
        }

        [Test]
        public void CFShiftedOnColumnDelete()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("CFShift");
                ws.Range("A1:A1").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AirForceBlue);
                ws.Range("A2:B2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AliceBlue);
                ws.Range("A3:C3").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Alizarin);
                ws.Range("B4:B6").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Almond);
                ws.Range("C7:D7").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Amaranth);
                ws.Cells("A1:D7").Value = 1;

                ws.Column(2).Delete();
                var cf = ws.ConditionalFormats.ToArray();

                Assert.AreEqual(4, cf.Length);
                Assert.AreEqual("A1:A1", cf[0].Range.RangeAddress.ToString());
                Assert.AreEqual("A2:A2", cf[1].Range.RangeAddress.ToString());
                Assert.AreEqual("A3:B3", cf[2].Range.RangeAddress.ToString());
                Assert.AreEqual("B7:C7", cf[3].Range.RangeAddress.ToString());
            }
        }

        [Test]
        public void CFShiftedOnRowDelete()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("CFShift");
                ws.Range("A1:A1").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AirForceBlue);
                ws.Range("B1:B2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AliceBlue);
                ws.Range("C1:C3").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Alizarin);
                ws.Range("D2:F2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Almond);
                ws.Range("G4:G5").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Amaranth);
                ws.Cells("A1:G5").Value = 1;

                ws.Row(2).Delete();
                var cf = ws.ConditionalFormats.ToArray();

                Assert.AreEqual(4, cf.Length);
                Assert.AreEqual("A1:A1", cf[0].Range.RangeAddress.ToString());
                Assert.AreEqual("B1:B1", cf[1].Range.RangeAddress.ToString());
                Assert.AreEqual("C1:C2", cf[2].Range.RangeAddress.ToString());
                Assert.AreEqual("G3:G4", cf[3].Range.RangeAddress.ToString());
            }
        }

        [Test]
        public void CFShiftedTruncateRange()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("CFShift");
                ws.AsRange().AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Red);
                var cf = ws.ConditionalFormats.Single();

                ws.Row(2).InsertRowsAbove(1);
                Assert.IsTrue(cf.Range.RangeAddress.IsValid);
                Assert.AreEqual($"1:{XLHelper.MaxRowNumber}", cf.Range.RangeAddress.ToString());

                ws.Column(2).InsertColumnsAfter(1);
                Assert.IsTrue(cf.Range.RangeAddress.IsValid);
                Assert.AreEqual($"1:{XLHelper.MaxRowNumber}", cf.Range.RangeAddress.ToString());
            }
        }
    }
}
