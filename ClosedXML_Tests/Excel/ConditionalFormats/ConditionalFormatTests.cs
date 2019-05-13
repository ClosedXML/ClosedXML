using ClosedXML.Excel;
using NUnit.Framework;
using System.IO;

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

            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\ConditionalFormattingOptional\inputfile_multiple_ranges.xlsx")))
            using (var ms = new MemoryStream())
            {
                TestHelper.CreateAndCompare(() =>
                {
                    var wb = new XLWorkbook(stream);
                    wb.SaveAs(ms,options);
                    return wb;
                }, @"Other\StyleReferenceFiles\ConditionalFormattingOptional\ConditionalFormattingOptional_not_consolidated.xlsx", false, options);
            }
        }

        [Test]
        public void OptionalRangeConsolidationWithConsolidation()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\ConditionalFormattingOptional\inputfile_multiple_ranges.xlsx")))
            using (var ms = new MemoryStream())
            {
                TestHelper.CreateAndCompare(() =>
                {
                    var wb = new XLWorkbook(stream);
                    wb.SaveAs(ms);
                    return wb;
                }, @"Other\StyleReferenceFiles\ConditionalFormattingOptional\ConditionalFormattingOptional_consolidated.xlsx");
            }
        }
    }
}
