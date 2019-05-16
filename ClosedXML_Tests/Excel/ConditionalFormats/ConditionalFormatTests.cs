using ClosedXML.Excel;
using NUnit.Framework;
using System.IO;
using System.Linq;

namespace ClosedXML_Tests.Excel.ConditionalFormats
{
    [TestFixture]
    public class ConditionalFormatTests
    {
        [Test]
        public void MaintainConditionalFormattingOrder()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\ConditionalFormattingOrder\inputfile.xlsx")))
            using (var ms = new MemoryStream())
            {
                TestHelper.CreateAndCompare(() =>
                {
                    var wb = new XLWorkbook(stream);
                    wb.SaveAs(ms);
                    return wb;
                }, @"Other\StyleReferenceFiles\ConditionalFormattingOrder\ConditionalFormattingOrder.xlsx");
            }
        }

        [TestCase(true, 7)]
        [TestCase(false, 8)]
        public void SaveOptionAffectsConsolidationConditionalFormatRanges(bool consolidateConditionalFormatRanges, int expectedCount)
        {
            var options = new SaveOptions
            {
                ConsolidateConditionalFormatRanges = consolidateConditionalFormatRanges
            };

            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");

            ws.Range("D2:D3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue();
            ws.Range("B2:B3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue();
            ws.Range("E2:E6").AddConditionalFormat().ColorScale().LowestValue(XLColor.Red).HighestValue(XLColor.Blue);
            ws.Range("F2:F6").AddConditionalFormat().ColorScale().LowestValue(XLColor.Red).HighestValue(XLColor.Blue);
            ws.Range("G2:G7").AddConditionalFormat().WhenIsUnique().Fill.SetBackgroundColor(XLColor.Blue);
            ws.Range("H2:H7").AddConditionalFormat().WhenIsUnique().Fill.SetBackgroundColor(XLColor.Blue);
            ws.Range("I2:I6").AddConditionalFormat().WhenContains("test");
            ws.Range("J2:J6").AddConditionalFormat().WhenContains("test");
            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms, options);
                var wb_saved = new XLWorkbook(ms);
                Assert.AreEqual(expectedCount, wb_saved.Worksheet("Sheet").ConditionalFormats.Count());
            }
        }

        [TestCase(true, 1)]
        [TestCase(false, 2)]
        public void SaveOptionAffectsConsolidationDataValidationRanges(bool consolidateDataValidationRanges, int expectedCount)
        {
            var options = new SaveOptions
            {
                ConsolidateDataValidationRanges = consolidateDataValidationRanges
            };

            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Range("C2:C5").SetDataValidation().Decimal.Between(1, 5);
            ws.Range("D2:D5").SetDataValidation().Decimal.Between(1, 5);

            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms, options);
                var wb_saved = new XLWorkbook(ms);
                Assert.AreEqual(expectedCount, wb_saved.Worksheet("Sheet").DataValidations.Count());
            }
        }
    }
}
