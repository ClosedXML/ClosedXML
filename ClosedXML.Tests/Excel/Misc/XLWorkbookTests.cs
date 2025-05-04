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
            Assert.That(cell, Is.Null);
        }

        [Test]
        public void Cell2()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
            IXLCell cell = wb.Cell("Sheet1!Result");
            Assert.That(cell, Is.Not.Null);
            Assert.That(cell.Value, Is.EqualTo(1));
        }

        [Test]
        public void Cell3()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result");
            IXLCell cell = wb.Cell("Sheet1!Result");
            Assert.That(cell, Is.Not.Null);
            Assert.That(cell.Value, Is.EqualTo(1));
        }

        [Test]
        public void Cells1()
        {
            var wb = new XLWorkbook();
            IXLCells cells = wb.Cells("ABC");
            Assert.That(cells, Is.Not.Null);
            Assert.That(cells.Count(), Is.EqualTo(0));
        }

        [Test]
        public void Cells2()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
            IXLCells cells = wb.Cells("Sheet1!Result, ABC");
            Assert.That(cells, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(cells.Count(), Is.EqualTo(1));
                Assert.That(cells.First().Value, Is.EqualTo(1));
            });
        }

        [Test]
        public void Cells3()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result");
            IXLCells cells = wb.Cells("Sheet1!Result, ABC");
            Assert.That(cells, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(cells.Count(), Is.EqualTo(1));
                Assert.That(cells.First().Value, Is.EqualTo(1));
            });
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

            Assert.Multiple(() =>
            {
                Assert.That(c1_full, Is.EqualTo(c1));
                Assert.That(c2_full, Is.EqualTo(c2));
            });
            Assert.Multiple(() =>
            {
                Assert.That(c1_full, Is.Not.Null);
                Assert.That(c2_full, Is.Not.Null);
            });
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

            Assert.That(c, Is.Null);
        }

        [Test]
        public void GetRangeFromFullAddress()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            var r1 = ws.Range("C123:D125");

            var r2 = wb.Range("Sheet1!C123:D125");

            Assert.That(r2, Is.SameAs(r1));
            Assert.That(r2, Is.Not.Null);
        }

        [TestCase("Sheet2!C1:D2")]
        [TestCase("Sheet1!A")]
        public void GetRangeFromNonExistingFullAddress(string rangeAddress)
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            var r = wb.Range(rangeAddress);

            Assert.That(r, Is.Null);
        }

        [Test]
        public void GetRangesFromFullAddress()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            var r1 = ws.Ranges("A1:B2,C1:E3");

            var r2 = wb.Ranges("Sheet1!A1:B2,Sheet1!C1:E3");

            Assert.That(r2, Has.Count.EqualTo(2));
            Assert.Multiple(() =>
            {
                Assert.That(r2.First(), Is.SameAs(r1.First()));
                Assert.That(r2.Last(), Is.SameAs(r1.Last()));
            });
        }

        [TestCase("Sheet2!C1:D2,Sheet2!F1:G4")]
        [TestCase("Sheet1!A,Sheet1!B")]
        public void GetRangesFromNonExistingFullAddress(string rangesAddress)
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            var r = wb.Ranges(rangesAddress);

            Assert.That(r, Is.Not.Null);
            Assert.False(r.Any());
        }

        [Test]
        public void Non_existent_defined_name_returns_null()
        {
            var wb = new XLWorkbook();
            var definedName = wb.DefinedName("ABC");
            Assert.That(definedName, Is.Null);
        }

        [Test]
        public void Sheet_specified_defined_name_is_retrieved_from_sheet_if_defined_there()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
            var definedName = wb.DefinedName("Sheet1!Result");
            Assert.That(definedName, Is.Not.Null);
            Assert.That(definedName.Ranges, Has.Count.EqualTo(1));
            Assert.Multiple(() =>
            {
                Assert.That(definedName.Ranges.Cells().Count(), Is.EqualTo(1));
                Assert.That(definedName.Ranges.First().FirstCell().Value, Is.EqualTo(1));
            });
        }

        [Test]
        public void Sheet_specified_defined_name_returns_null_if_not_defined_in_sheet_nor_workbook()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var definedName = wb.DefinedName("Sheet1!Result");
            Assert.That(definedName, Is.Null);
        }

        [Test]
        public void Sheet_specified_defined_name_falls_back_to_workbook_scoped_defined_name_if_not_defined_in_sheet()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result");
            var definedName = wb.DefinedName("Sheet1!Result");
            Assert.That(definedName, Is.Not.Null);
            Assert.That(definedName.Ranges, Has.Count.EqualTo(1));
            Assert.Multiple(() =>
            {
                Assert.That(definedName.Ranges.Cells().Count(), Is.EqualTo(1));
                Assert.That(definedName.Ranges.First().FirstCell().Value, Is.EqualTo(1));
            });
        }

        [Test]
        public void Range1()
        {
            var wb = new XLWorkbook();
            IXLRange range = wb.Range("ABC");
            Assert.That(range, Is.Null);
        }

        [Test]
        public void Range2()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
            IXLRange range = wb.Range("Sheet1!Result");
            Assert.That(range, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(range.Cells().Count(), Is.EqualTo(1));
                Assert.That(range.FirstCell().Value, Is.EqualTo(1));
            });
        }

        [Test]
        public void Range3()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result");
            IXLRange range = wb.Range("Sheet1!Result");
            Assert.That(range, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(range.Cells().Count(), Is.EqualTo(1));
                Assert.That(range.FirstCell().Value, Is.EqualTo(1));
            });
        }

        [Test]
        public void Ranges1()
        {
            var wb = new XLWorkbook();
            IXLRanges ranges = wb.Ranges("ABC");
            Assert.That(ranges, Is.Not.Null);
            Assert.That(ranges, Is.Empty);
        }

        [Test]
        public void Ranges2()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
            IXLRanges ranges = wb.Ranges("Sheet1!Result, ABC");
            Assert.That(ranges, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(ranges.Cells().Count(), Is.EqualTo(1));
                Assert.That(ranges.First().FirstCell().Value, Is.EqualTo(1));
            });
        }

        [Test]
        public void Ranges3()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result");
            IXLRanges ranges = wb.Ranges("Sheet1!Result, ABC");
            Assert.That(ranges, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(ranges.Cells().Count(), Is.EqualTo(1));
                Assert.That(ranges.First().FirstCell().Value, Is.EqualTo(1));
            });
        }

        [Test]
        public void WbNamedCell()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell");
            Assert.Multiple(() =>
            {
                Assert.That(wb.Cell("TestCell").GetText(), Is.EqualTo("Test"));
                Assert.That(ws.Cell("TestCell").GetText(), Is.EqualTo("Test"));
            });
        }

        [Test]
        public void WbNamedCells()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell");
            ws.Cell(2, 1).SetValue("B").AddToNamed("Test2");
            IXLCells wbCells = wb.Cells("TestCell, Test2");
            Assert.Multiple(() =>
            {
                Assert.That(wbCells.First().GetText(), Is.EqualTo("Test"));
                Assert.That(wbCells.Last().GetText(), Is.EqualTo("B"));
            });

            IXLCells wsCells = ws.Cells("TestCell, Test2");
            Assert.Multiple(() =>
            {
                Assert.That(wsCells.First().GetText(), Is.EqualTo("Test"));
                Assert.That(wsCells.Last().GetText(), Is.EqualTo("B"));
            });
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
            Assert.That(wb.Range("TestRange").RangeAddress.ToString(), Is.EqualTo(original.RangeAddress.ToStringFixed()));
            Assert.That(ws.Range("TestRange").RangeAddress.ToString(), Is.EqualTo(original.RangeAddress.ToStringFixed()));
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
            Assert.Multiple(() =>
            {
                Assert.That(wbRanges.First().RangeAddress.ToString(), Is.EqualTo(original.RangeAddress.ToStringFixed()));
                Assert.That(wbRanges.Last().RangeAddress.ToStringFixed(), Is.EqualTo("$A$3:$A$3"));
            });

            IXLRanges wsRanges = wb.Ranges("TestRange, Test2");
            Assert.Multiple(() =>
            {
                Assert.That(wsRanges.First().RangeAddress.ToString(), Is.EqualTo(original.RangeAddress.ToStringFixed()));
                Assert.That(wsRanges.Last().RangeAddress.ToStringFixed(), Is.EqualTo("$A$3:$A$3"));
            });
        }

        [Test]
        public void WbNamedRangesOneString()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            wb.DefinedNames.Add("TestRange", "Sheet1!$A$1,Sheet1!$A$3");

            IXLRanges wbRanges = ws.Ranges("TestRange");
            Assert.Multiple(() =>
            {
                Assert.That(wbRanges.First().RangeAddress.ToStringFixed(), Is.EqualTo("$A$1:$A$1"));
                Assert.That(wbRanges.Last().RangeAddress.ToStringFixed(), Is.EqualTo("$A$3:$A$3"));
            });

            IXLRanges wsRanges = ws.Ranges("TestRange");
            Assert.Multiple(() =>
            {
                Assert.That(wsRanges.First().RangeAddress.ToStringFixed(), Is.EqualTo("$A$1:$A$1"));
                Assert.That(wsRanges.Last().RangeAddress.ToStringFixed(), Is.EqualTo("$A$3:$A$3"));
            });
        }

        [Test]
        public void WbProtect1()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            wb.Protect();
            Assert.Multiple(() =>
            {
                Assert.That(wb.LockStructure, Is.True);
                Assert.That(wb.LockWindows, Is.False);
                Assert.That(wb.IsPasswordProtected, Is.False);
            });
        }

        [Test]
        public void WbProtect2()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            wb.Protect(XLWorkbookProtectionElements.Windows);
            Assert.Multiple(() =>
            {
                Assert.That(wb.LockStructure, Is.True);
                Assert.That(wb.LockWindows, Is.False);
                Assert.That(wb.IsPasswordProtected, Is.False);
            });
        }

        [Test]
        public void WbProtect3()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            wb.Protect("Abc@123");
            Assert.Multiple(() =>
            {
                Assert.That(wb.LockStructure, Is.True);
                Assert.That(wb.LockWindows, Is.False);
                Assert.That(wb.IsPasswordProtected, Is.True);
            });
            Assert.Throws<InvalidOperationException>(() => wb.Protect());
            Assert.Throws<InvalidOperationException>(() => wb.Unprotect());
            Assert.Throws<ArgumentException>(() => wb.Unprotect("Cde@345"));
        }

        [Test]
        public void WbProtect4()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            wb.Protect();
            Assert.Multiple(() =>
            {
                Assert.That(wb.LockStructure, Is.True);
                Assert.That(wb.LockWindows, Is.False);
                Assert.That(wb.IsPasswordProtected, Is.False);
            });
            wb.Unprotect();
            wb.Protect("Abc@123");
            Assert.Multiple(() =>
            {
                Assert.That(wb.LockStructure, Is.True);
                Assert.That(wb.LockWindows, Is.False);
                Assert.That(wb.IsPasswordProtected, Is.True);
            });
        }

        [Test]
        public void WbProtect5()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            wb.Protect("Abc@123", XLProtectionAlgorithm.DefaultProtectionAlgorithm, XLWorkbookProtectionElements.Windows);
            Assert.Multiple(() =>
            {
                Assert.That(wb.LockStructure, Is.True);
                Assert.That(wb.LockWindows, Is.False);
                Assert.That(wb.IsPasswordProtected, Is.True);
            });
            wb.Unprotect("Abc@123");
            Assert.Multiple(() =>
            {
                Assert.That(wb.LockStructure, Is.False);
                Assert.That(wb.LockWindows, Is.False);
                Assert.That(wb.IsPasswordProtected, Is.False);
            });
        }

        [Test]
        public void FileSharingProperties()
        {
            using var ms = new MemoryStream();
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
                Assert.Multiple(() =>
                {
                    Assert.That(wb.FileSharing.ReadOnlyRecommended, Is.True);
                    Assert.That(wb.FileSharing.UserName, Is.EqualTo(Environment.UserName));
                });
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
