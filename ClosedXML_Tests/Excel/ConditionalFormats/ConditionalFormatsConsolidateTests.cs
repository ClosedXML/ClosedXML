using System.IO;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel.ConditionalFormats
{
    [TestFixture]
    public class ConditionalFormatsConsolidateTests
    {
        [Test]
        public void ConsecutivelyRowsConsolidateTest()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("B2:C2").AddConditionalFormat());
            SetFormat1(ws.Range("B4:C4").AddConditionalFormat());
            SetFormat1(ws.Range("B3:C3").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            var format = ws.ConditionalFormats.First();
            Assert.AreEqual("B2:C4", format.Range.RangeAddress.ToStringRelative());
            Assert.AreEqual("F2", format.Values.Values.First().Value);
        }

        [Test]
        public void ConsecutivelyColumnsConsolidateTest()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("D2:D3").AddConditionalFormat());
            SetFormat1(ws.Range("B2:B3").AddConditionalFormat());
            SetFormat1(ws.Range("C2:C3").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            var format = ws.ConditionalFormats.First();
            Assert.AreEqual("B2:D3", format.Ranges.First().RangeAddress.ToStringRelative());
            Assert.AreEqual("F2", format.Values.Values.First().Value);
        }

        [Test]
        public void Contains1ConsolidateTest()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");
            
            SetFormat1(ws.Range("B11:D12").AddConditionalFormat());
            SetFormat1(ws.Range("C12:D12").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            var format = ws.ConditionalFormats.First();
            Assert.AreEqual("B11:D12", format.Range.RangeAddress.ToStringRelative());
            Assert.AreEqual("F11", format.Values.Values.First().Value);
        }

        [Test]
        public void Contains2ConsolidateTest()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("B14:C14").AddConditionalFormat());
            SetFormat1(ws.Range("B14:B14").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            var format = ws.ConditionalFormats.First();
            Assert.AreEqual("B14:C14", format.Range.RangeAddress.ToStringRelative());
            Assert.AreEqual("F14", format.Values.Values.First().Value);
        }

        [Test]
        public void SuperimposedConsolidateTest()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("B16:D18").AddConditionalFormat());
            SetFormat1(ws.Range("B18:D19").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            var format = ws.ConditionalFormats.First();
            Assert.AreEqual("B16:D19", format.Range.RangeAddress.ToStringRelative());
            Assert.AreEqual("F16", format.Values.Values.First().Value);
        }

        [Test]
        public void DifferentFormatNoConsolidateTest()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("B11:D12").AddConditionalFormat());
            SetFormat2(ws.Range("C12:D12").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.AreEqual(2, ws.ConditionalFormats.Count());
        }

        [Test]
        public void ConsolidatePreservesPriorities()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("A1:A5").AddConditionalFormat());
            SetFormat2(ws.Range("A1:A5").AddConditionalFormat());
            SetFormat2(ws.Range("A6:A10").AddConditionalFormat());
            SetFormat1(ws.Range("A6:A10").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.AreEqual(3, ws.ConditionalFormats.Count());
            Assert.AreEqual((ws.ConditionalFormats.First().Style as XLStyle).Value, (ws.ConditionalFormats.Last().Style as XLStyle).Value);
            Assert.AreNotEqual((ws.ConditionalFormats.First().Style as XLStyle).Value, (ws.ConditionalFormats.ElementAt(1).Style as XLStyle).Value);
        }


        [Test]
        public void ConsolidatePreservesPriorities2()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("A1:A1").AddConditionalFormat());
            SetFormat2(ws.Range("A2:A3").AddConditionalFormat());
            SetFormat1(ws.Range("A2:A6").AddConditionalFormat());
            SetFormat1(ws.Range("A7:A8").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.AreEqual(3, ws.ConditionalFormats.Count());
            Assert.AreEqual((ws.ConditionalFormats.First().Style as XLStyle).Value, (ws.ConditionalFormats.Last().Style as XLStyle).Value);
            Assert.AreNotEqual((ws.ConditionalFormats.First().Style as XLStyle).Value, (ws.ConditionalFormats.ElementAt(1).Style as XLStyle).Value);
            Assert.IsTrue(ws.ConditionalFormats.All(cf => cf.Ranges.Count == 1), "Number of ranges in consolidated conditional formats is expected to be 1");
            Assert.AreEqual("A1:A1", ws.ConditionalFormats.ElementAt(0).Ranges.Single().RangeAddress.ToString());
            Assert.AreEqual("A2:A3", ws.ConditionalFormats.ElementAt(1).Ranges.Single().RangeAddress.ToString());
            Assert.AreEqual("A2:A8", ws.ConditionalFormats.ElementAt(2).Ranges.Single().RangeAddress.ToString());
        }


        [Test]
        public void ConsolidateShiftsFormulaRelativelyToTopMostCell()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            var ranges = ws.Ranges("B3:B8,C3:C4,A3:A4,C5:C8,A5:A8").Cast<XLRange>();
            var cf = new XLConditionalFormat(ranges);
            cf.Values.Add(new XLFormula("=A3=$D3"));
            cf.Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.ConditionalFormats.Add(cf);

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            Assert.AreEqual((ws.ConditionalFormats.Single().Style as XLStyle).Value, (cf.Style as XLStyle).Value);
            Assert.AreEqual("A3:C8", ws.ConditionalFormats.Single().Ranges.Single().RangeAddress.ToString());
            Assert.IsTrue(ws.ConditionalFormats.Single().Values.Single().Value.IsFormula);
            Assert.AreEqual("A3=$D3", ws.ConditionalFormats.Single().Values.Single().Value.Value);
        }

        [Test]
        public void ColorScaleComparing()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet");

                var ranges = ws.Ranges("B3:B8,C3:C4,A3:A4,C5:C8,A5:A8").Cast<XLRange>();
                var cf1 = new XLConditionalFormat(ranges);
                cf1.ColorScale()
                    .LowestValue(XLColor.Red)
                    .HighestValue(XLColor.Green);

                var cf2 = new XLConditionalFormat(ranges);
                cf2.ColorScale()
                    .LowestValue(XLColor.Red)
                    .HighestValue(XLColor.Green);
                Assert.True(XLConditionalFormat.NoRangeComparer.Equals(cf1, cf2));
            }
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
