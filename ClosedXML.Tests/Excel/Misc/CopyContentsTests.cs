using System;
using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.Misc
{
    [TestFixture]
    public class CopyContentsTests
    {
        private static void CopyRowAsRange(IXLWorksheet originalSheet, int originalRowNumber, IXLWorksheet destSheet,
            int destRowNumber)
        {
            {
                IXLRow destinationRow = destSheet.Row(destRowNumber);
                destinationRow.Clear();

                IXLRow originalRow = originalSheet.Row(originalRowNumber);
                int columnNumber = originalRow.LastCellUsed(XLCellsUsedOptions.All).Address.ColumnNumber;

                IXLRange originalRange = originalSheet.Range(originalRowNumber, 1, originalRowNumber, columnNumber);
                IXLRange destRange = destSheet.Range(destRowNumber, 1, destRowNumber, columnNumber);
                originalRange.CopyTo(destRange);
            }
        }

        [Test]
        public void CopyConditionalFormatsCount()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().AddConditionalFormat().WhenContains("1").Fill.SetBackgroundColor(XLColor.Blue);
            ws.Cell("A2").CopyFrom(ws.FirstCell().AsRange());
            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(2));
        }

        [Test]
        public void CopyConditionalFormatsFixedNum()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "1";
            ws.Cell("B1").Value = "1";
            ws.Cell("A1").AddConditionalFormat().WhenEquals(1).Fill.SetBackgroundColor(XLColor.Blue);
            ws.Cell("A2").CopyFrom(ws.Cell("A1").AsRange());
            Assert.That(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)), Is.True);
            Assert.That(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)), Is.True);
        }

        [Test]
        public void CopyConditionalFormatsFixedString()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "A";
            ws.Cell("B1").Value = "B";
            ws.Cell("A1").AddConditionalFormat().WhenEquals("A").Fill.SetBackgroundColor(XLColor.Blue);
            ws.Cell("A2").CopyFrom(ws.Cell("A1").AsRange());
            Assert.That(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "A" && !v.Value.IsFormula)), Is.True);
            Assert.That(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "A" && !v.Value.IsFormula)), Is.True);
        }

        [Test]
        public void CopyConditionalFormatsFixedStringNum()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "1";
            ws.Cell("B1").Value = "1";
            ws.Cell("A1").AddConditionalFormat().WhenEquals("1").Fill.SetBackgroundColor(XLColor.Blue);
            ws.Cell("A2").CopyFrom(ws.Cell("A1").AsRange());
            Assert.That(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)), Is.True);
            Assert.That(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)), Is.True);
        }

        [Test]
        public void CopyConditionalFormatsRelative()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "1";
            ws.Cell("B1").Value = "1";
            ws.Cell("A1").AddConditionalFormat().WhenEquals("=B1").Fill.SetBackgroundColor(XLColor.Blue);
            ws.Cell("A2").CopyFrom(ws.Cell("A1").AsRange());
            Assert.Multiple(() =>
            {
                Assert.That(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "B1" && v.Value.IsFormula)), Is.True);
                Assert.That(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "B2" && v.Value.IsFormula)), Is.True);
            });
        }

        [Test]
        public void TestRowCopyContents()
        {
            var workbook = new XLWorkbook();
            IXLWorksheet originalSheet = workbook.Worksheets.Add("original");
            IXLWorksheet copyRowSheet = workbook.Worksheets.Add("copy row");
            IXLWorksheet copyRowAsRangeSheet = workbook.Worksheets.Add("copy row as range");
            IXLWorksheet copyRangeSheet = workbook.Worksheets.Add("copy range");

            originalSheet.Cell("A2").SetValue("test value");
            originalSheet.Range("A2:E2").Merge();

            {
                IXLRange originalRange = originalSheet.Range("A2:E2");
                IXLRange destinationRange = copyRangeSheet.Range("A2:E2");

                originalRange.CopyTo(destinationRange);
            }
            CopyRowAsRange(originalSheet, 2, copyRowAsRangeSheet, 3);
            {
                IXLRow originalRow = originalSheet.Row(2);
                IXLRow destinationRow = copyRowSheet.Row(2);
                copyRowSheet.Cell("G2").Value = "must be removed after copy";
                originalRow.CopyTo(destinationRow);
            }
            TestHelper.SaveWorkbook(workbook, "Misc", "CopyRowContents.xlsx");
        }

        [Test]
        public void UpdateCellsWorksheetTest()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            ws1.Cell(1, 1).Value = "hello, world.";

            var ws2 = ws1.CopyTo("Sheet2");

            Assert.Multiple(() =>
            {
                Assert.That(ws1.FirstCell().Address.Worksheet.Name, Is.EqualTo("Sheet1"));
                Assert.That(ws2.FirstCell().Address.Worksheet.Name, Is.EqualTo("Sheet2"));
            });
        }

        [Test]
        public void CopyHyperlinksAmongSheets()
        {
            using var wb = new XLWorkbook();
            var source = wb.AddWorksheet();
            var target = wb.AddWorksheet();
            source.Cell("A1")
                .SetValue("link")
                .CreateHyperlink()
                .SetValues("https://example.com", "Test tooltip");

            source.Cell("A1").AsRange().CopyTo(target.Cell("B7"));

            var cell = target.Cell("B7");
            Assert.Multiple(() =>
            {
                Assert.That(cell.HasHyperlink, Is.True);
                Assert.That(cell.GetHyperlink().IsExternal, Is.True);
            });
            Assert.Multiple(() =>
            {
                Assert.That(cell.GetHyperlink().ExternalAddress, Is.EqualTo(new Uri("https://example.com")));
                Assert.That(cell.GetHyperlink().Tooltip, Is.EqualTo("Test tooltip"));
            });
        }
    }
}
