using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML_Tests.Excel
{
    /// <summary>
    ///     Summary description for UnitTest1
    /// </summary>
    [TestFixture]
    public class NamedRangesTests
    {
        [Test]
        public void MovingRanges()
        {
            var wb = new XLWorkbook();

            IXLWorksheet sheet1 = wb.Worksheets.Add("Sheet1");
            IXLWorksheet sheet2 = wb.Worksheets.Add("Sheet2");

            wb.NamedRanges.Add("wbNamedRange",
                "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
            sheet1.NamedRanges.Add("sheet1NamedRange",
                "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
            sheet2.NamedRanges.Add("sheet2NamedRange", "Sheet1!A1,Sheet2!A1");

            sheet1.Row(1).InsertRowsAbove(2);
            sheet1.Row(1).Delete();
            sheet1.Column(1).InsertColumnsBefore(2);
            sheet1.Column(1).Delete();

            Assert.AreEqual("Sheet1!$C$3,Sheet1!$C$4:$D$4,Sheet2!$D$3:$D$4,Sheet1!$7:$8,Sheet1!$G:$H",
                wb.NamedRanges.First().RefersTo);
            Assert.AreEqual("Sheet1!$C$3,Sheet1!$C$4:$D$4,Sheet2!$D$3:$D$4,Sheet1!$7:$8,Sheet1!$G:$H",
                sheet1.NamedRanges.First().RefersTo);
            Assert.AreEqual("Sheet1!B2,Sheet2!A1", sheet2.NamedRanges.First().RefersTo);
        }

        [Test]
        public void WbContainsWsNamedRange()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().AddToNamed("Name", XLScope.Worksheet);

            Assert.IsTrue(wb.NamedRanges.Contains("Sheet1!Name"));
            Assert.IsFalse(wb.NamedRanges.Contains("Sheet1!NameX"));

            Assert.IsNotNull(wb.NamedRange("Sheet1!Name"));
            Assert.IsNull(wb.NamedRange("Sheet1!NameX"));

            Boolean result1 = wb.NamedRanges.TryGetValue("Sheet1!Name", out IXLNamedRange range1);
            Assert.IsTrue(result1);
            Assert.IsNotNull(range1);

            Boolean result2 = wb.NamedRanges.TryGetValue("Sheet1!NameX", out IXLNamedRange range2);
            Assert.IsFalse(result2);
            Assert.IsNull(range2);
        }

        [Test]
        public void WorkbookContainsNamedRange()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().AddToNamed("Name");

            Assert.IsTrue(wb.NamedRanges.Contains("Name"));
            Assert.IsFalse(wb.NamedRanges.Contains("NameX"));

            Assert.IsNotNull(wb.NamedRange("Name"));
            Assert.IsNull(wb.NamedRange("NameX"));

            Boolean result1 = wb.NamedRanges.TryGetValue("Name", out IXLNamedRange range1);
            Assert.IsTrue(result1);
            Assert.IsNotNull(range1);

            Boolean result2 = wb.NamedRanges.TryGetValue("NameX", out IXLNamedRange range2);
            Assert.IsFalse(result2);
            Assert.IsNull(range2);
        }

        [Test]
        public void WorksheetContainsNamedRange()
        {
            IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");
            ws.FirstCell().AddToNamed("Name", XLScope.Worksheet);

            Assert.IsTrue(ws.NamedRanges.Contains("Name"));
            Assert.IsFalse(ws.NamedRanges.Contains("NameX"));

            Assert.IsNotNull(ws.NamedRange("Name"));
            Assert.IsNull(ws.NamedRange("NameX"));

            Boolean result1 = ws.NamedRanges.TryGetValue("Name", out IXLNamedRange range1);
            Assert.IsTrue(result1);
            Assert.IsNotNull(range1);

            Boolean result2 = ws.NamedRanges.TryGetValue("NameX", out IXLNamedRange range2);
            Assert.IsFalse(result2);
            Assert.IsNull(range2);
        }

        [Test]
        public void DeleteColumnUsedInNamedRange()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.FirstCell().SetValue("Column1");
                ws.FirstCell().CellRight().SetValue("Column2").Style.Font.SetBold();
                ws.FirstCell().CellRight(2).SetValue("Column3");
                ws.NamedRanges.Add("MyRange", "A1:C1");

                ws.Column(1).Delete();

                Assert.IsTrue(ws.Cell("A1").Style.Font.Bold);
                Assert.AreEqual("Column3", ws.Cell("B1").GetValue<string>());
                Assert.IsEmpty(ws.Cell("C1").GetValue<string>());
            }
        }

        [Test]
        public void TestInvalidNamedRangeOnWorkbookScope()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.FirstCell().SetValue("Column1");
                ws.FirstCell().CellRight().SetValue("Column2").Style.Font.SetBold();
                ws.FirstCell().CellRight(2).SetValue("Column3");

                Assert.Throws<ArgumentException>(() => wb.NamedRanges.Add("MyRange", "A1:C1"));
            }
        }

        [Test]
        public void NamedRangesWhenCopyingWorksheets()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet("Sheet1");
                ws1.FirstCell().Value = Enumerable.Range(1, 10);
                wb.NamedRanges.Add("wbNamedRange", ws1.Range("A1:A10"));
                ws1.NamedRanges.Add("wsNamedRange", ws1.Range("A3"));

                var ws2 = wb.AddWorksheet("Sheet2");
                ws2.FirstCell().Value = Enumerable.Range(101, 10);
                ws1.NamedRanges.Add("wsNamedRangeAcrossSheets", ws2.Range("A4"));

                ws1.Cell("C1").FormulaA1 = "=wbNamedRange";
                ws1.Cell("C2").FormulaA1 = "=wsNamedRange";
                ws1.Cell("C3").FormulaA1 = "=wsNamedRangeAcrossSheets";

                Assert.AreEqual(1, ws1.Cell("C1").Value);
                Assert.AreEqual(3, ws1.Cell("C2").Value);
                Assert.AreEqual(104, ws1.Cell("C3").Value);

                var wsCopy = ws1.CopyTo("Copy");
                Assert.AreEqual(1, wsCopy.Cell("C1").Value);
                Assert.AreEqual(3, wsCopy.Cell("C2").Value);
                Assert.AreEqual(104, wsCopy.Cell("C3").Value);

                Assert.AreEqual("Sheet1!A1:A10", wb.NamedRange("wbNamedRange").Ranges.First().RangeAddress.ToStringRelative(true));
                Assert.AreEqual("Copy!A3:A3", wsCopy.NamedRange("wsNamedRange").Ranges.First().RangeAddress.ToStringRelative(true));
                Assert.AreEqual("Sheet2!A4:A4", wsCopy.NamedRange("wsNamedRangeAcrossSheets").Ranges.First().RangeAddress.ToStringRelative(true));
            }
        }
    }
}
