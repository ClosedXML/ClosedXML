using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML_Tests.Excel.Misc
{
    /// <summary>
    ///     Summary description for UnitTest1
    /// </summary>
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
            ws.Cell("A2").Value = ws.FirstCell();
            Assert.AreEqual(2, ws.ConditionalFormats.Count());
        }

        [Test]
        public void CopyConditionalFormatsFixedNum()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "1";
            ws.Cell("B1").Value = "1";
            ws.Cell("A1").AddConditionalFormat().WhenEquals(1).Fill.SetBackgroundColor(XLColor.Blue);
            ws.Cell("A2").Value = ws.Cell("A1");
            Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)));
            Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)));
        }

        [Test]
        public void CopyConditionalFormatsFixedString()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "A";
            ws.Cell("B1").Value = "B";
            ws.Cell("A1").AddConditionalFormat().WhenEquals("A").Fill.SetBackgroundColor(XLColor.Blue);
            ws.Cell("A2").Value = ws.Cell("A1");
            Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "A" && !v.Value.IsFormula)));
            Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "A" && !v.Value.IsFormula)));
        }

        [Test]
        public void CopyConditionalFormatsFixedStringNum()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "1";
            ws.Cell("B1").Value = "1";
            ws.Cell("A1").AddConditionalFormat().WhenEquals("1").Fill.SetBackgroundColor(XLColor.Blue);
            ws.Cell("A2").Value = ws.Cell("A1");
            Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)));
            Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)));
        }

        [Test]
        public void CopyConditionalFormatsRelative()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "1";
            ws.Cell("B1").Value = "1";
            ws.Cell("A1").AddConditionalFormat().WhenEquals("=B1").Fill.SetBackgroundColor(XLColor.Blue);
            ws.Cell("A2").Value = ws.Cell("A1");
            Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "B1" && v.Value.IsFormula)));
            Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "B2" && v.Value.IsFormula)));
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
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.Worksheets.Add("Sheet1");
                ws1.Cell(1, 1).Value = "hello, world.";

                var ws2 = ws1.CopyTo("Sheet2");

                Assert.AreEqual("Sheet1", ws1.FirstCell().Address.Worksheet.Name);
                Assert.AreEqual("Sheet2", ws2.FirstCell().Address.Worksheet.Name);
            }
        }
    }
}
