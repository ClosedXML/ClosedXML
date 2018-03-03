using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class XLWorksheetTests
    {
        [Test]
        public void ColumnCountTime()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            DateTime start = DateTime.Now;
            ws.ColumnCount();
            DateTime end = DateTime.Now;
            Assert.IsTrue((end - start).TotalMilliseconds < 500);
        }

        [Test]
        public void CopyConditionalFormatsCount()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().AddConditionalFormat().WhenContains("1").Fill.SetBackgroundColor(XLColor.Blue);
            IXLWorksheet ws2 = ws.CopyTo("Sheet2");
            Assert.AreEqual(1, ws2.ConditionalFormats.Count());
        }

        [Test]
        public void DeletingSheets1()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet3");
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets.Add("Sheet1", 1);

            wb.Worksheet("Sheet3").Delete();

            Assert.AreEqual("Sheet1", wb.Worksheet(1).Name);
            Assert.AreEqual("Sheet2", wb.Worksheet(2).Name);
            Assert.AreEqual(2, wb.Worksheets.Count);
        }

        [Test]
        public void InsertingSheets1()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet1");
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets.Add("Sheet3");

            Assert.AreEqual("Sheet1", wb.Worksheet(1).Name);
            Assert.AreEqual("Sheet2", wb.Worksheet(2).Name);
            Assert.AreEqual("Sheet3", wb.Worksheet(3).Name);
        }

        [Test]
        public void InsertingSheets2()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets.Add("Sheet1", 1);
            wb.Worksheets.Add("Sheet3");

            Assert.AreEqual("Sheet1", wb.Worksheet(1).Name);
            Assert.AreEqual("Sheet2", wb.Worksheet(2).Name);
            Assert.AreEqual("Sheet3", wb.Worksheet(3).Name);
        }

        [Test]
        public void InsertingSheets3()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet3");
            wb.Worksheets.Add("Sheet2", 1);
            wb.Worksheets.Add("Sheet1", 1);

            Assert.AreEqual("Sheet1", wb.Worksheet(1).Name);
            Assert.AreEqual("Sheet2", wb.Worksheet(2).Name);
            Assert.AreEqual("Sheet3", wb.Worksheet(3).Name);
        }

        [Test]
        public void AddingDuplicateSheetNameThrowsException()
        {
            using (var wb = new XLWorkbook())
            {
                IXLWorksheet ws;
                ws = wb.AddWorksheet("Sheet1");

                Assert.Throws<ArgumentException>(() => wb.AddWorksheet("Sheet1"));

                //Sheet names are case insensitive
                Assert.Throws<ArgumentException>(() => wb.AddWorksheet("sheet1"));
            }
        }

        [Test]
        public void MergedRanges()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("A1:B2").Merge();
            ws.Range("C1:D3").Merge();
            ws.Range("D2:E2").Merge();

            Assert.AreEqual(2, ws.MergedRanges.Count);
            Assert.AreEqual("A1:B2", ws.MergedRanges.First().RangeAddress.ToStringRelative());
            Assert.AreEqual("D2:E2", ws.MergedRanges.Last().RangeAddress.ToStringRelative());
        }

        [Test]
        public void RowCountTime()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            DateTime start = DateTime.Now;
            ws.RowCount();
            DateTime end = DateTime.Now;
            Assert.IsTrue((end - start).TotalMilliseconds < 500);
        }

        [Test]
        public void SheetsWithCommas()
        {
            using (var wb = new XLWorkbook())
            {
                var sourceSheetName = "Sheet1, Sheet3";
                var ws = wb.Worksheets.Add(sourceSheetName);
                ws.Cell("A1").Value = 1;
                ws.Cell("A2").Value = 2;
                ws.Cell("B2").Value = 3;

                ws = wb.Worksheets.Add("Formula");
                ws.FirstCell().FormulaA1 = string.Format("=SUM('{0}'!A1:A2,'{0}'!B1:B2)", sourceSheetName);

                var value = ws.FirstCell().Value;
                Assert.AreEqual(6, value);
            }
        }

        [Test]
        public void CanRenameWorksheet()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet("Sheet1");
                var ws2 = wb.AddWorksheet("Sheet2");

                ws1.Name = "New sheet name";
                Assert.AreEqual("New sheet name", ws1.Name);

                ws2.Name = "sheet2";
                Assert.AreEqual("sheet2", ws2.Name);

                Assert.Throws<ArgumentException>(() => ws1.Name = "SHEET2");
            }
        }

        [Test]
        public void TryGetWorksheet()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet("Sheet1");
                var ws2 = wb.AddWorksheet("Sheet2");

                IXLWorksheet ws;
                Assert.IsTrue(wb.Worksheets.TryGetWorksheet("Sheet1", out ws));
                Assert.IsTrue(wb.Worksheets.TryGetWorksheet("sheet1", out ws));
                Assert.IsTrue(wb.Worksheets.TryGetWorksheet("sHEeT1", out ws));
                Assert.IsFalse(wb.Worksheets.TryGetWorksheet("Sheeeet2", out ws));

                Assert.IsTrue(wb.TryGetWorksheet("Sheet1", out ws));
                Assert.IsTrue(wb.TryGetWorksheet("sheet1", out ws));
                Assert.IsTrue(wb.TryGetWorksheet("sHEeT1", out ws));
                Assert.IsFalse(wb.TryGetWorksheet("Sheeeet2", out ws));
            }
        }

        [Test]
        public void HideWorksheet()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    wb.Worksheets.Add("VisibleSheet");
                    wb.Worksheets.Add("HiddenSheet").Hide();
                    wb.SaveAs(ms);
                }

                // unhide the hidden sheet
                using (var wb = new XLWorkbook(ms))
                {
                    Assert.AreEqual(XLWorksheetVisibility.Visible, wb.Worksheet("VisibleSheet").Visibility);
                    Assert.AreEqual(XLWorksheetVisibility.Hidden, wb.Worksheet("HiddenSheet").Visibility);

                    var ws = wb.Worksheet("HiddenSheet");
                    ws.Unhide().Name = "NoAlsoVisible";

                    Assert.AreEqual(XLWorksheetVisibility.Visible, ws.Visibility);

                    wb.Save();
                }

                using (var wb = new XLWorkbook(ms))
                {
                    Assert.AreEqual(XLWorksheetVisibility.Visible, wb.Worksheet("VisibleSheet").Visibility);
                    Assert.AreEqual(XLWorksheetVisibility.Visible, wb.Worksheet("NoAlsoVisible").Visibility);
                }
            }
        }

        [Test]
        public void CanCopySheetsWithAllAnchorTypes()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                ws.CopyTo("Copy1");

                var ws2 = wb.Worksheets.Skip(1).First();
                ws2.CopyTo("Copy2");

                var ws3 = wb.Worksheets.Skip(2).First();
                ws3.CopyTo("Copy3");

                var ws4 = wb.Worksheets.Skip(3).First();
                ws3.CopyTo("Copy4");
            }
        }

        [Test]
        public void WorksheetNameCannotStartWithApostrophe()
        {
            var title = "'StartsWithApostrophe";
            TestDelegate addWorksheet = () =>
            {
                using (var wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(title);
                }
            };

            Assert.Throws(typeof(ArgumentException), addWorksheet);
        }

        [Test]
        public void WorksheetNameCannotEndWithApostrophe()
        {
            var title = "EndsWithApostrophe'";
            TestDelegate addWorksheet = () =>
            {
                using (var wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(title);
                }
            };

            Assert.Throws(typeof(ArgumentException), addWorksheet);
        }

        [Test]
        public void WorksheetNameCanContainApostrophe()
        {
            var title = "With'Apostrophe";
            var savedTitle = "";
            TestDelegate saveAndOpenWorkbook = () =>
            {
                using (var ms = new MemoryStream())
                {
                    using (var wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add(title);
                        wb.Worksheets.First().Cell(1, 1).FormulaA1 = $"{title}!A2";
                        wb.SaveAs(ms);
                    }

                    using (var wb = new XLWorkbook(ms))
                    {
                        savedTitle = wb.Worksheets.First().Name;
                    }
                }
            };

            Assert.DoesNotThrow(saveAndOpenWorkbook);
            Assert.AreEqual(title, savedTitle);
        }
    }
}
