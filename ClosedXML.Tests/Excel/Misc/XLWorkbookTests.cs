using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class XLWorkbookTests
    {
        [Test]
        public void Cell1()
        {
            var wb = new XLWorkbook();
            IXLCell cell = wb.Cell("ABC");
            Assert.IsNull(cell);
        }

        [Test]
        public void Cell2()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
            IXLCell cell = wb.Cell("Sheet1!Result");
            Assert.IsNotNull(cell);
            Assert.AreEqual(1, cell.GetValue<Int32>());
        }

        [Test]
        public void Cell3()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result");
            IXLCell cell = wb.Cell("Sheet1!Result");
            Assert.IsNotNull(cell);
            Assert.AreEqual(1, cell.GetValue<Int32>());
        }

        [Test]
        public void Cells1()
        {
            var wb = new XLWorkbook();
            IXLCells cells = wb.Cells("ABC");
            Assert.IsNotNull(cells);
            Assert.AreEqual(0, cells.Count());
        }

        [Test]
        public void Cells2()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
            IXLCells cells = wb.Cells("Sheet1!Result, ABC");
            Assert.IsNotNull(cells);
            Assert.AreEqual(1, cells.Count());
            Assert.AreEqual(1, cells.First().GetValue<Int32>());
        }

        [Test]
        public void Cells3()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result");
            IXLCells cells = wb.Cells("Sheet1!Result, ABC");
            Assert.IsNotNull(cells);
            Assert.AreEqual(1, cells.Count());
            Assert.AreEqual(1, cells.First().GetValue<Int32>());
        }

        [Test]
        public void GetCellFromFullAddress()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLWorksheet ws2 = wb.AddWorksheet("O'Sheet 2");
            var c1 = ws.Cell("C123");
            var c2 = ws2.Cell("B7");

            var c1_full = wb.Cell("Sheet1!C123");
            var c2_full = wb.Cell("'O'Sheet 2'!B7");

            Assert.AreSame(c1, c1_full);
            Assert.AreSame(c2, c2_full);
            Assert.NotNull(c1_full);
            Assert.NotNull(c2_full);
        }

        [TestCase("Sheet1")]
        [TestCase("Sheet1!")]
        [TestCase("Sheet2!")]
        [TestCase("Sheet2!C1")]
        [TestCase("Sheet1!ZZZ1")]
        [TestCase("Sheet1!A")]
        public void GetCellFromNonExistingFullAddress(string address)
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            var c = wb.Cell(address);

            Assert.IsNull(c);
        }

        [Test]
        public void GetRangeFromFullAddress()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            var r1 = ws.Range("C123:D125");

            var r2 = wb.Range("Sheet1!C123:D125");

            Assert.AreSame(r1, r2);
            Assert.NotNull(r2);
        }

        [TestCase("Sheet2!C1:D2")]
        [TestCase("Sheet1!A")]
        public void GetRangeFromNonExistingFullAddress(string rangeAddress)
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            var r = wb.Range(rangeAddress);

            Assert.IsNull(r);
        }

        [Test]
        public void GetRangesFromFullAddress()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            var r1 = ws.Ranges("A1:B2,C1:E3");

            var r2 = wb.Ranges("Sheet1!A1:B2,Sheet1!C1:E3");

            Assert.AreEqual(2, r2.Count);
            Assert.AreSame(r1.First(), r2.First());
            Assert.AreSame(r1.Last(), r2.Last());
        }

        [TestCase("Sheet2!C1:D2,Sheet2!F1:G4")]
        [TestCase("Sheet1!A,Sheet1!B")]
        public void GetRangesFromNonExistingFullAddress(string rangesAddress)
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            var r = wb.Ranges(rangesAddress);

            Assert.NotNull(r);
            Assert.False(r.Any());
        }

        [Test]
        public void NamedRange1()
        {
            var wb = new XLWorkbook();
            IXLNamedRange range = wb.NamedRange("ABC");
            Assert.IsNull(range);
        }

        [Test]
        public void NamedRange2()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
            IXLNamedRange range = wb.NamedRange("Sheet1!Result");
            Assert.IsNotNull(range);
            Assert.AreEqual(1, range.Ranges.Count);
            Assert.AreEqual(1, range.Ranges.Cells().Count());
            Assert.AreEqual(1, range.Ranges.First().FirstCell().GetValue<Int32>());
        }

        [Test]
        public void NamedRange3()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLNamedRange range = wb.NamedRange("Sheet1!Result");
            Assert.IsNull(range);
        }

        [Test]
        public void NamedRange4()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result");
            IXLNamedRange range = wb.NamedRange("Sheet1!Result");
            Assert.IsNotNull(range);
            Assert.AreEqual(1, range.Ranges.Count);
            Assert.AreEqual(1, range.Ranges.Cells().Count());
            Assert.AreEqual(1, range.Ranges.First().FirstCell().GetValue<Int32>());
        }

        [Test]
        public void Range1()
        {
            var wb = new XLWorkbook();
            IXLRange range = wb.Range("ABC");
            Assert.IsNull(range);
        }

        [Test]
        public void Range2()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
            IXLRange range = wb.Range("Sheet1!Result");
            Assert.IsNotNull(range);
            Assert.AreEqual(1, range.Cells().Count());
            Assert.AreEqual(1, range.FirstCell().GetValue<Int32>());
        }

        [Test]
        public void Range3()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result");
            IXLRange range = wb.Range("Sheet1!Result");
            Assert.IsNotNull(range);
            Assert.AreEqual(1, range.Cells().Count());
            Assert.AreEqual(1, range.FirstCell().GetValue<Int32>());
        }

        [Test]
        public void Ranges1()
        {
            var wb = new XLWorkbook();
            IXLRanges ranges = wb.Ranges("ABC");
            Assert.IsNotNull(ranges);
            Assert.AreEqual(0, ranges.Count());
        }

        [Test]
        public void Ranges2()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
            IXLRanges ranges = wb.Ranges("Sheet1!Result, ABC");
            Assert.IsNotNull(ranges);
            Assert.AreEqual(1, ranges.Cells().Count());
            Assert.AreEqual(1, ranges.First().FirstCell().GetValue<Int32>());
        }

        [Test]
        public void Ranges3()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result");
            IXLRanges ranges = wb.Ranges("Sheet1!Result, ABC");
            Assert.IsNotNull(ranges);
            Assert.AreEqual(1, ranges.Cells().Count());
            Assert.AreEqual(1, ranges.First().FirstCell().GetValue<Int32>());
        }

        [Test]
        public void WbNamedCell()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell");
            Assert.AreEqual("Test", wb.Cell("TestCell").GetString());
            Assert.AreEqual("Test", ws.Cell("TestCell").GetString());
        }

        [Test]
        public void WbNamedCells()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell");
            ws.Cell(2, 1).SetValue("B").AddToNamed("Test2");
            IXLCells wbCells = wb.Cells("TestCell, Test2");
            Assert.AreEqual("Test", wbCells.First().GetString());
            Assert.AreEqual("B", wbCells.Last().GetString());

            IXLCells wsCells = ws.Cells("TestCell, Test2");
            Assert.AreEqual("Test", wsCells.First().GetString());
            Assert.AreEqual("B", wsCells.Last().GetString());
        }

        [Test]
        public void WbNamedRange()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("A");
            ws.Cell(2, 1).SetValue("B");
            IXLRange original = ws.Range("A1:A2");
            original.AddToNamed("TestRange");
            Assert.AreEqual(original.RangeAddress.ToStringFixed(), wb.Range("TestRange").RangeAddress.ToString());
            Assert.AreEqual(original.RangeAddress.ToStringFixed(), ws.Range("TestRange").RangeAddress.ToString());
        }

        [Test]
        public void WbNamedRanges()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("A");
            ws.Cell(2, 1).SetValue("B");
            ws.Cell(3, 1).SetValue("C").AddToNamed("Test2");
            IXLRange original = ws.Range("A1:A2");
            original.AddToNamed("TestRange");
            IXLRanges wbRanges = wb.Ranges("TestRange, Test2");
            Assert.AreEqual(original.RangeAddress.ToStringFixed(), wbRanges.First().RangeAddress.ToString());
            Assert.AreEqual("$A$3:$A$3", wbRanges.Last().RangeAddress.ToStringFixed());

            IXLRanges wsRanges = wb.Ranges("TestRange, Test2");
            Assert.AreEqual(original.RangeAddress.ToStringFixed(), wsRanges.First().RangeAddress.ToString());
            Assert.AreEqual("$A$3:$A$3", wsRanges.Last().RangeAddress.ToStringFixed());
        }

        [Test]
        public void WbNamedRangesOneString()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            wb.NamedRanges.Add("TestRange", "Sheet1!$A$1,Sheet1!$A$3");

            IXLRanges wbRanges = ws.Ranges("TestRange");
            Assert.AreEqual("$A$1:$A$1", wbRanges.First().RangeAddress.ToStringFixed());
            Assert.AreEqual("$A$3:$A$3", wbRanges.Last().RangeAddress.ToStringFixed());

            IXLRanges wsRanges = ws.Ranges("TestRange");
            Assert.AreEqual("$A$1:$A$1", wsRanges.First().RangeAddress.ToStringFixed());
            Assert.AreEqual("$A$3:$A$3", wsRanges.Last().RangeAddress.ToStringFixed());
        }

        [Test]
        public void WbProtect1()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                wb.Protect();
                Assert.IsTrue(wb.LockStructure);
                Assert.IsFalse(wb.LockWindows);
                Assert.IsFalse(wb.IsPasswordProtected);
            }
        }

        [Test]
        public void WbProtect2()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                wb.Protect(true, false);
                Assert.IsTrue(wb.LockStructure);
                Assert.IsFalse(wb.LockWindows);
                Assert.IsFalse(wb.IsPasswordProtected);
            }
        }

        [Test]
        public void WbProtect3()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                wb.Protect("Abc@123");
                Assert.IsTrue(wb.LockStructure);
                Assert.IsFalse(wb.LockWindows);
                Assert.IsTrue(wb.IsPasswordProtected);
                Assert.Throws<InvalidOperationException>(() => wb.Protect());
                Assert.Throws<InvalidOperationException>(() => wb.Unprotect());
                Assert.Throws<ArgumentException>(() => wb.Unprotect("Cde@345"));
            }
        }

        [Test]
        public void WbProtect4()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                wb.Protect();
                Assert.IsTrue(wb.LockStructure);
                Assert.IsFalse(wb.LockWindows);
                Assert.IsFalse(wb.IsPasswordProtected);
                wb.Unprotect();
                wb.Protect("Abc@123");
                Assert.IsTrue(wb.LockStructure);
                Assert.IsFalse(wb.LockWindows);
                Assert.IsTrue(wb.IsPasswordProtected);
            }
        }

        [Test]
        public void WbProtect5()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                wb.Protect(true, false, "Abc@123");
                Assert.IsTrue(wb.LockStructure);
                Assert.IsFalse(wb.LockWindows);
                Assert.IsTrue(wb.IsPasswordProtected);
                wb.Unprotect("Abc@123");
                Assert.IsFalse(wb.LockStructure);
                Assert.IsFalse(wb.LockWindows);
                Assert.IsFalse(wb.IsPasswordProtected);
            }
        }

        [Test]
        public void FileSharingProperties()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    wb.AddWorksheet("Sheet1").Cell("A1").Value = "Hello world!";
                    wb.FileSharing.ReadOnlyRecommended = true;
                    wb.FileSharing.UserName = Environment.UserName;
                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    Assert.IsTrue(wb.FileSharing.ReadOnlyRecommended);
                    Assert.AreEqual(Environment.UserName, wb.FileSharing.UserName);
                }
            }
        }

        [Test]
        public void AccessDisposedWorkbookThrowsException()
        {
            IXLWorkbook wb;
            using (wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                ws.FirstCell().SetValue("Hello world");
            }

            Assert.Throws<ObjectDisposedException>(() => Console.WriteLine(wb.Worksheets.First().FirstCell().Value));
        }
    }
}
