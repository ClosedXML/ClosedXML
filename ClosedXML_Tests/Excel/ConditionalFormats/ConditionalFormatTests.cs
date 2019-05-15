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

        [Test]
        public void OptionalRangeConsolidationWithoutConsolidation()
        {
            var options = new SaveOptions
            {
                ConsolidateConditionalFormatRanges = false,
                ConsolidateDataValidationRanges = false
            };

            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Range("D2:D3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue();
            ws.Range("B2:B3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue();
            ws.Range("C2:C5").SetDataValidation().Decimal.Between(1, 5);
            ws.Range("D2:D5").SetDataValidation().Decimal.Between(1, 5);
            ws.Range("E2:E6").AddConditionalFormat().ColorScale().LowestValue(XLColor.Red).HighestValue(XLColor.Blue);
            ws.Range("F2:F6").AddConditionalFormat().ColorScale().LowestValue(XLColor.Red).HighestValue(XLColor.Blue);

            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms, options);
                var wb_saved = new XLWorkbook(ms);
                Assert.AreEqual(6, wb_saved.Worksheet("Sheet").ConditionalFormats.Count() + wb_saved.Worksheet("Sheet").DataValidations.Count());
            }
        }

        [Test]
        public void OptionalRangeConsolidationWithConsolidation()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Range("D2:D3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue();
            ws.Range("B2:B3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue();
            ws.Range("C2:C5").SetDataValidation().Decimal.Between(1, 5);
            ws.Range("D2:D5").SetDataValidation().Decimal.Between(1, 5);
            ws.Range("E2:E6").AddConditionalFormat().ColorScale().LowestValue(XLColor.Red).HighestValue(XLColor.Blue);
            ws.Range("F2:F6").AddConditionalFormat().ColorScale().LowestValue(XLColor.Red).HighestValue(XLColor.Blue);

            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms);
                var wb_saved = new XLWorkbook(ms);
                Assert.AreEqual(3, wb_saved.Worksheet("Sheet").ConditionalFormats.Count() + wb_saved.Worksheet("Sheet").DataValidations.Count());
            }
        }
    }
}
