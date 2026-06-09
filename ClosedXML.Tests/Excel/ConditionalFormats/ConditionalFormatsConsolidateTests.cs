using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.ConditionalFormats
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
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add();

            // Format2 A1:A5 and A6:A10 can be consolidated without changing a priority if other CF rules.
            // Format1 A1:A5 and A6:A10 can't be consolidated, because there is a CF rule between them.
            SetFormat1(ws.Range("A1:A5").AddConditionalFormat());
            SetFormat2(ws.Range("A1:A5").AddConditionalFormat());
            SetFormat2(ws.Range("A6:A10").AddConditionalFormat());
            SetFormat1(ws.Range("A6:A10").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.AreEqual(3, ws.ConditionalFormats.Count());
            var cf1 = ws.ConditionalFormats.First();
            var cf2 = ws.ConditionalFormats.ElementAt(1);
            var cf3 = ws.ConditionalFormats.Last();
            Assert.That(cf1, Is.EqualTo(cf3).Using(new CfFormatComaparer()));
            Assert.That(cf1, Is.Not.EqualTo(cf2).Using(new CfFormatComaparer()));
        }

        [Test]
        public void ConsolidatePreservesPriorities2()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add();

            SetFormat1(ws.Range("A1:A1").AddConditionalFormat());
            SetFormat2(ws.Range("A2:A3").AddConditionalFormat());
            SetFormat1(ws.Range("A2:A6").AddConditionalFormat());
            SetFormat1(ws.Range("A7:A8").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.AreEqual(3, ws.ConditionalFormats.Count());
            var cf1 = ws.ConditionalFormats.ElementAt(0);
            var cf2 = ws.ConditionalFormats.ElementAt(1);
            var cf3 = ws.ConditionalFormats.ElementAt(2);
            Assert.That(cf1, Is.EqualTo(cf3).Using(new CfFormatComaparer()));
            Assert.That(cf1, Is.Not.EqualTo(cf2).Using(new CfFormatComaparer()));
            Assert.IsTrue(ws.ConditionalFormats.All(cf => cf.Ranges.Count() == 1), "Number of ranges in consolidated conditional formats is expected to be 1");
            Assert.AreEqual("A1:A1", cf1.Ranges.Single().RangeAddress.ToString());
            Assert.AreEqual("A2:A3", cf2.Ranges.Single().RangeAddress.ToString());
            Assert.AreEqual("A2:A8", cf3.Ranges.Single().RangeAddress.ToString());
        }

        [Test]
        public void ConsolidateShiftsFormulaRelativelyToTopMostCell()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add();

            var ranges = ws.Ranges("B3:B8,C3:C4,A3:A4,C5:C8,A5:A8");
            var cf = ranges.First().AddConditionalFormat();
            cf.Ranges = ranges;
            cf.Values.Add(new XLFormula("=A3=$D3"));
            cf.Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.ConditionalFormats.Add(cf);

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.AreEqual(1, ws.ConditionalFormats.Count());
            var consolidatedCf = ws.ConditionalFormats.Single();
            Assert.That(consolidatedCf, Is.EqualTo(cf).Using (new CfFormatComaparer()));
            Assert.AreEqual("A3:C8", consolidatedCf.Ranges.Single().RangeAddress.ToString());
            Assert.IsTrue(consolidatedCf.Values.Single().Value.IsFormula);
            Assert.AreEqual("A3=$D3", consolidatedCf.Values.Single().Value.Value);
        }

        [Test]
        public void ColorScaleComparing()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet");

                var ranges = ws.Ranges("B3:B8,C3:C4,A3:A4,C5:C8,A5:A8");
                var cf1 = ranges.First().AddConditionalFormat();
                cf1.Ranges = ranges;
                cf1.ColorScale()
                    .LowestValue(XLColor.Red)
                    .HighestValue(XLColor.Green);

                var cf2 = ranges.First().AddConditionalFormat();
                cf2.Ranges = ranges;
                cf2.ColorScale()
                    .LowestValue(XLColor.Red)
                    .HighestValue(XLColor.Green);
                Assert.AreNotSame(cf1, cf2);
                Assert.True(XLConditionalFormat.NoRangeComparer.Equals((XLConditionalFormat)cf1, (XLConditionalFormat)cf2));
            }
        }

        private static void SetFormat1(IXLConditionalFormat format)
        {
            format.WhenEquals("=" + format.Range.FirstCell().CellRight(4).Address.ToStringRelative()).Fill.SetBackgroundColor(XLColor.Blue);
        }

        private static void SetFormat2(IXLConditionalFormat format)
        {
            format.WhenEquals(5).Fill.SetBackgroundColor(XLColor.AliceBlue);
        }

        private class CfFormatComaparer : IEqualityComparer<IXLConditionalFormat>
        {
            public bool Equals(IXLConditionalFormat x, IXLConditionalFormat y)
            {
                var lhs = (XLConditionalFormat)x;
                var rhs = (XLConditionalFormat)y;
                return lhs.FormatValue == rhs.FormatValue;
            }

            public int GetHashCode([DisallowNull] IXLConditionalFormat obj)
            {
                return HashCode.Combine(((XLConditionalFormat)obj).FormatValue);
            }
        }
    }
}
