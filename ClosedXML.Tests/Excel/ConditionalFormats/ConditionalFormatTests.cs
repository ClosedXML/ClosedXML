using ClosedXML.Excel;
using NUnit.Framework;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;

namespace ClosedXML.Tests.Excel.ConditionalFormats
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
            ws.Range("C2:C5").CreateDataValidation().Decimal.Between(1, 5);
            ws.Range("D2:D5").CreateDataValidation().Decimal.Between(1, 5);

            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms, options);
                var wb_saved = new XLWorkbook(ms);
                Assert.AreEqual(expectedCount, wb_saved.Worksheet("Sheet").DataValidations.Count());
            }
        }

        [TestCase("en-US")]
        [TestCase("fr-FR")]
        [TestCase("ru-RU")]
        public void SaveConditionalFormat_CultureIndependent(string culture)
        {
            using (var ms = new MemoryStream())
            {
                var expectedValue = 1.5;
                Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo(culture);
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet();
                    var i = 1;
                    ws.Cell(i++, 1).AddConditionalFormat().WhenEquals(expectedValue).Fill.SetBackgroundColor(XLColor.Red);
                    ws.Cell(i++, 1).AddConditionalFormat().WhenNotEquals(expectedValue).Fill.SetBackgroundColor(XLColor.Red);
                    ws.Cell(i++, 1).AddConditionalFormat().WhenGreaterThan(expectedValue).Fill.SetBackgroundColor(XLColor.Red);
                    ws.Cell(i++, 1).AddConditionalFormat().WhenLessThan(expectedValue).Fill.SetBackgroundColor(XLColor.Red);
                    ws.Cell(i++, 1).AddConditionalFormat().WhenEqualOrGreaterThan(expectedValue).Fill.SetBackgroundColor(XLColor.Red);
                    ws.Cell(i++, 1).AddConditionalFormat().WhenEqualOrLessThan(expectedValue).Fill.SetBackgroundColor(XLColor.Red);
                    ws.Cell(i++, 1).AddConditionalFormat().WhenBetween(expectedValue, expectedValue).Fill.SetBackgroundColor(XLColor.Red);
                    ws.Cell(i++, 1).AddConditionalFormat().WhenNotBetween(expectedValue, expectedValue).Fill.SetBackgroundColor(XLColor.Red);

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheets.First();

                    var conditionalFormatValues = ws.ConditionalFormats
                        .SelectMany(cf => cf.Values.Values)
                        .Select(v => v.Value)
                        .Distinct();

                    Assert.AreEqual(1, conditionalFormatValues.Count());
                    Assert.AreEqual("1.5", conditionalFormatValues.Single());
                }
            }
        }

        [Test]
        public void CellIs_type_reads_only_required_formula_arguments()
        {
            // The CellIs uses formula tags as arguments. Some producers generate extra empty
            // formula tags and ClosedXml should be able to load CellIs conditional formatting
            // with such extra tags without an exception. The test file has been modified to
            // include extra formula tags and test checks that extra tags are ignored.
            TestHelper.LoadAndAssert((_, ws) =>
            {
                AssertFormulaArgs(ws, XLCFOperator.Between, "$D$2", "$E$2");
                AssertFormulaArgs(ws, XLCFOperator.NotBetween, "$D$3", "$E$3");
                AssertFormulaArgs(ws, XLCFOperator.GreaterThan, "$D$4");
                AssertFormulaArgs(ws, XLCFOperator.LessThan, "$D$5");
                AssertFormulaArgs(ws, XLCFOperator.Equal, "$D$6");
            }, @"Other\ConditionalFormats\Extra_formulas_CellIs_type.xlsx");

            static void AssertFormulaArgs(IXLWorksheet ws, XLCFOperator cfOperator, params string[] expectedFormulas)
            {
                var cf = ws.ConditionalFormats.Single(cf => cf.ConditionalFormatType == XLConditionalFormatType.CellIs && cf.Operator == cfOperator);
                Assert.AreEqual(expectedFormulas.Length, cf.Values.Count);
                CollectionAssert.AreEqual(expectedFormulas, cf.Values.Select(v => v.Value.Value));
            }
        }

        [Test]
        public void Expression_type_skips_empty_formula_tags()
        {
            // The Expression uses formula tag as arguments. Some producers generate extra empty
            // formula tags and ClosedXml should be able to load Expression conditional formatting
            // with such extra tags without an exception. The test file has been modified to
            // include extra formula tags and test checks that extra tags are ignored.
            TestHelper.LoadAndAssert((_, ws) =>
            {
                AssertFormulaArgs(ws, "A1:A1", "$C$1=5");
                AssertFormulaArgs(ws, "A2:A2", "$C$2=4");
            }, @"Other\ConditionalFormats\Extra_formulas_Expression_type.xlsx");

            static void AssertFormulaArgs(IXLWorksheet ws, string range, string expectedFormula)
            {
                var cf = ws.ConditionalFormats.Single(cf => cf.ConditionalFormatType == XLConditionalFormatType.Expression && cf.Range.RangeAddress.ToString() == range);
                Assert.AreEqual(1, cf.Values.Count);
                CollectionAssert.AreEqual(expectedFormula, cf.Values[1].Value);
            }
        }
    }
}
