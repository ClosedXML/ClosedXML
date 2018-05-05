using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel.Drawings;

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

            Assert.AreEqual("A1:B2", ws.Cell("A2").MergedRange().RangeAddress.ToStringRelative());
            Assert.AreEqual("D2:E2", ws.Cell("D2").MergedRange().RangeAddress.ToStringRelative());

            Assert.AreEqual(null, ws.Cell("Z10").MergedRange());
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
                wb.AddWorksheet("Sheet1");
                wb.AddWorksheet("Sheet2");

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

        [Test]
        public void CopyWorksheetPreservesContents()
        {
            using (var wb1 = new XLWorkbook())
            using (var wb2 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");

                ws1.Cell("A1").Value = "A1 value";
                ws1.Cell("A2").Value = 100;
                ws1.Cell("D4").Value = new DateTime(2018, 5, 1);

                var ws2 = ws1.CopyTo(wb2, "Copy");

                Assert.AreEqual("A1 value", ws2.Cell("A1").Value);
                Assert.AreEqual(100, ws2.Cell("A2").Value);
                Assert.AreEqual(new DateTime(2018, 5, 1), ws2.Cell("D4").Value);
            }
        }

        [Test]
        public void CopyWorksheetPreservesFormulae()
        {
            using (var wb1 = new XLWorkbook())
            using (var wb2 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");

                ws1.Cell("A1").FormulaA1 = "10*10";
                ws1.Cell("A2").FormulaA1 = "A1 * 2";

                var ws2 = ws1.CopyTo(wb2, "Copy");

                Assert.AreEqual("10*10", ws2.Cell("A1").FormulaA1);
                Assert.AreEqual("A1 * 2", ws2.Cell("A2").FormulaA1);
            }
        }
        
        [Test]
        public void CopyWorksheetPreservesRowHeightsAfterSave()
        {
            using (var ms = new MemoryStream())
            using (var wb1 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");
                using (var wb2 = new XLWorkbook())
                {
                    ws1.RowHeight = 55;
                    ws1.Row(2).Height = 0;
                    ws1.Row(3).Height = 20;

                    var ws2 = ws1.CopyTo(wb2, "Copy");

                    Assert.AreEqual(ws1.RowHeight, ws2.RowHeight);
                    for (int i = 1; i <= 3; i++)
                    {
                        Assert.AreEqual(ws1.Row(i).Height, ws2.Row(i).Height);
                    }

                    wb2.SaveAs(ms);
                }

                using (var wb2 = new XLWorkbook(ms))
                {
                    var ws2 = wb2.Worksheets.First();

                    Assert.AreEqual(ws1.RowHeight, ws2.RowHeight);
                    for (int i = 1; i <= 3; i++)
                    {
                        Assert.AreEqual(ws1.Row(i).Height, ws2.Row(i).Height);
                    }
                }
            }
        }

        [Test]
        public void CopyWorksheetPreservesColumnWidthsAfterSave()
        {
            using (var ms = new MemoryStream())
            using (var wb1 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");
                using (var wb2 = new XLWorkbook())
                {
                    ws1.ColumnWidth = 160;
                    ws1.Column(2).Width = 0;
                    ws1.Column(3).Width = 240;

                    var ws2 = ws1.CopyTo(wb2, "Copy");

                    Assert.AreEqual(ws1.ColumnWidth, ws2.ColumnWidth);
                    for (int i = 1; i <= 3; i++)
                    {
                        Assert.AreEqual(ws1.Column(i).Width, ws2.Column(i).Width);
                    }

                    wb2.SaveAs(ms);
                }

                using (var wb2 = new XLWorkbook(ms))
                {
                    var ws2 = wb2.Worksheets.First();

                    Assert.AreEqual(ws1.ColumnWidth, ws2.ColumnWidth, 1);
                    for (int i = 1; i <= 3; i++)
                    {
                        Assert.AreEqual(ws1.Column(i).Width, ws2.Column(i).Width, 1);
                    }
                }
            }
        }

        [Test]
        public void CopyWorksheetPreservesMergedCells()
        {
            using (var wb1 = new XLWorkbook())
            using (var wb2 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");

                ws1.Range("A:A").Merge();
                ws1.Range("B1:C2").Merge();

                var ws2 = ws1.CopyTo(wb2, "Copy");

                Assert.AreEqual(ws1.MergedRanges.Count, ws2.MergedRanges.Count);
                for (int i = 0; i < ws1.MergedRanges.Count; i++)
                {
                    Assert.AreEqual(ws1.MergedRanges.ElementAt(i).RangeAddress.ToString(),
                                    ws2.MergedRanges.ElementAt(i).RangeAddress.ToString());
                }
            }
        }

        [Test]
        public void CopyWorksheetAcrossWorkbooksPreservesNamedRanges()
        {
            using (var wb1 = new XLWorkbook())
            using (var wb2 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");

                ws1.Range("A1:A2").AddToNamed("GLOBAL", XLScope.Workbook);
                ws1.Ranges("B1:B2,D1:D2").AddToNamed("LOCAL", XLScope.Worksheet);

                var ws2 = ws1.CopyTo(wb2, "Copy");

                Assert.AreEqual(ws1.NamedRanges.Count(), ws2.NamedRanges.Count());
                for (int i = 0; i < ws1.NamedRanges.Count(); i++)
                {
                    var nr1 = ws1.NamedRanges.ElementAt(i);
                    var nr2 = ws2.NamedRanges.ElementAt(i);
                    Assert.AreEqual(nr1.Ranges.ToString(), nr2.Ranges.ToString());
                    Assert.AreEqual(nr1.Scope, nr2.Scope);
                    Assert.AreEqual(nr1.Name, nr2.Name);
                    Assert.AreEqual(nr1.Visible, nr2.Visible);
                    Assert.AreEqual(nr1.Comment, nr2.Comment);
                }
            }
        }


        [Test]
        public void CopyWorksheeInsideWorkbookMakesNamedRangesLocal()
        {
            using (var wb1 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");

                ws1.Range("A1:A2").AddToNamed("GLOBAL", XLScope.Workbook);
                ws1.Ranges("B1:B2,D1:D2").AddToNamed("LOCAL", XLScope.Worksheet);

                var ws2 = ws1.CopyTo("Copy");

                Assert.AreEqual(ws1.NamedRanges.Count(), ws2.NamedRanges.Count());
                for (int i = 0; i < ws1.NamedRanges.Count(); i++)
                {
                    var nr1 = ws1.NamedRanges.ElementAt(i);
                    var nr2 = ws2.NamedRanges.ElementAt(i);

                    Assert.AreEqual(XLScope.Worksheet, nr2.Scope);

                    Assert.AreEqual(nr1.Ranges.ToString(), nr2.Ranges.ToString());
                    Assert.AreEqual(nr1.Name, nr2.Name);
                    Assert.AreEqual(nr1.Visible, nr2.Visible);
                    Assert.AreEqual(nr1.Comment, nr2.Comment);
                }
            }
        }

        [Test]
        public void CopyWorksheetPreservesStyles()
        {
            using (var ms = new MemoryStream())
            using (var wb1 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");

                ws1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws1.Range("A1:B2").Style.Font.FontSize = 25;
                ws1.Cell("C3").Style.Fill.BackgroundColor = XLColor.Red;
                ws1.Cell("C4").Style.Fill.BackgroundColor = XLColor.AliceBlue;
                ws1.Cell("C4").Value = "Non empty";

                using (var wb2 = new XLWorkbook())
                {
                    var ws2 = ws1.CopyTo(wb2, "Copy");
                    AssertStylesAreEqual(ws1, ws2);
                    wb2.SaveAs(ms);
                }

                using (var wb2 = new XLWorkbook(ms))
                {
                    var ws2 = wb2.Worksheet("Copy");
                    AssertStylesAreEqual(ws1, ws2);
                }
            }

            void AssertStylesAreEqual(IXLWorksheet ws1, IXLWorksheet ws2)
            {
                Assert.AreEqual((ws1.Style as XLStyle).Value, (ws2.Style as XLStyle).Value,
                    "Worksheet styles differ");
                var cellsUsed = ws1.Range(ws1.FirstCell(), ws1.LastCellUsed()).Cells();
                foreach (var cell in cellsUsed)
                {
                    var style1 = (cell.Style as XLStyle).Value;
                    var style2 = (ws2.Cell(cell.Address.ToString()).Style as XLStyle).Value;
                    Assert.AreEqual(style1, style2, $"Cell {cell.Address} styles differ");
                }
            }
        }

        [Test]
        public void CopyWorksheetPreservesConditionalFormats()
        {
            using (var wb1 = new XLWorkbook())
            using (var wb2 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");

                ws1.Range("A:A").AddConditionalFormat()
                    .WhenContains("0").Fill.SetBackgroundColor(XLColor.Red);
                ws1.Range("B1:C2").AddConditionalFormat()
                    .WhenEqualOrGreaterThan(100).Font.SetBold();

                var ws2 = ws1.CopyTo(wb2, "Copy");

                Assert.AreEqual(ws1.ConditionalFormats.Count(), ws2.ConditionalFormats.Count());
                for (int i = 0; i < ws1.ConditionalFormats.Count(); i++)
                {
                    Assert.AreEqual(ws1.ConditionalFormats.ElementAt(i).Ranges.ToString(),
                                    ws2.ConditionalFormats.ElementAt(i).Ranges.ToString());
                    Assert.AreEqual(ws1.ConditionalFormats.ElementAt(i).Style,
                                    ws2.ConditionalFormats.ElementAt(i).Style);
                    Assert.AreEqual(ws1.ConditionalFormats.ElementAt(i).Values.Single(),
                                    ws2.ConditionalFormats.ElementAt(i).Values.Single());
                }
            }
        }

        [Test]
        public void CopyWorksheetPreservesDataValidation()
        {
            using (var wb1 = new XLWorkbook())
            using (var wb2 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");

                var dv1 = ws1.Range("A:A").SetDataValidation();
                dv1.WholeNumber.EqualTo(2);
                dv1.ErrorStyle = XLErrorStyle.Warning;
                dv1.ErrorTitle = "Number out of range";
                dv1.ErrorMessage = "This cell only allows the number 2.";
                
                var dv2 = ws1.Ranges("B2:C3,D4:E5").SetDataValidation();
                dv2.Decimal.GreaterThan(5);
                dv2.ErrorStyle = XLErrorStyle.Stop;
                dv2.ErrorTitle = "Decimal number out of range";
                dv2.ErrorMessage = "This cell only allows decimals greater than 5.";

                var dv3 = ws1.Cell("D1").SetDataValidation();
                dv3.TextLength.EqualOrLessThan(10);
                dv3.ErrorStyle = XLErrorStyle.Information;
                dv3.ErrorTitle = "Text length out of range";
                dv3.ErrorMessage = "You entered more than 10 characters.";
                
                var ws2 = ws1.CopyTo(wb2, "Copy");

                Assert.AreEqual(ws1.DataValidations.Count(), ws2.DataValidations.Count());
                for (int i = 0; i < ws1.DataValidations.Count(); i++)
                {
                    var original = ws1.DataValidations.ElementAt(i);
                    var copy = ws2.DataValidations.ElementAt(i);

                    Assert.AreEqual(original.Ranges.ToString(), copy.Ranges.ToString());
                    Assert.AreEqual(original.AllowedValues, copy.AllowedValues);
                    Assert.AreEqual(original.Operator, copy.Operator);
                    Assert.AreEqual(original.ErrorStyle, copy.ErrorStyle);
                    Assert.AreEqual(original.ErrorTitle, copy.ErrorTitle);
                    Assert.AreEqual(original.ErrorMessage, copy.ErrorMessage);
                }
            }
        }

        [Test]
        public void CopyWorksheetPreservesPictures()
        {
            using (var ms = new MemoryStream())
            using (var resourceStream = Assembly.GetAssembly(typeof(ClosedXML_Examples.BasicTable))
                .GetManifestResourceStream("ClosedXML_Examples.Resources.SampleImage.jpg"))
            using (var bitmap = Bitmap.FromStream(resourceStream) as Bitmap)
            using (var wb1 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");

                var picture = ws1.AddPicture(bitmap, "MyPicture")
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .MoveTo(50, 50)
                    .WithSize(200, 200);

                using (var wb2 = new XLWorkbook())
                {
                    var ws2 = ws1.CopyTo(wb2, "Copy");
                    AssertPicturesAreEqual(ws1, ws2);
                    wb2.SaveAs(ms);
                }

                using (var wb2 = new XLWorkbook(ms))
                {
                    var ws2 = wb2.Worksheet("Copy");
                    AssertPicturesAreEqual(ws1, ws2);
                }
            }

            void AssertPicturesAreEqual(IXLWorksheet ws1, IXLWorksheet ws2)
            {
                Assert.AreEqual(ws1.Pictures.Count(), ws2.Pictures.Count());

                for (int i = 0; i < ws1.Pictures.Count(); i++)
                {
                    var original = ws1.Pictures.ElementAt(i);
                    var copy = ws2.Pictures.ElementAt(i);
                    Assert.AreEqual(ws2, copy.Worksheet);

                    Assert.AreEqual(original.Format, copy.Format);
                    Assert.AreEqual(original.Height, copy.Height);
                    Assert.AreEqual(original.Id, copy.Id);
                    Assert.AreEqual(original.Left, copy.Left);
                    Assert.AreEqual(original.Name, copy.Name);
                    Assert.AreEqual(original.Placement, copy.Placement);
                    Assert.AreEqual(original.Top, copy.Top);
                    Assert.AreEqual(original.TopLeftCellAddress.ToString(), copy.TopLeftCellAddress.ToString());
                    Assert.AreEqual(original.Width, copy.Width);
                    Assert.AreEqual(original.ImageStream.ToArray(), copy.ImageStream.ToArray(), "Image streams differ");
                }
            }
        }

        [Test, Ignore("Muted until #836 is fixed")]
        public void CopyWorksheetChangesAbsoluteReferencesInFormulae()
        {
            using (var wb1 = new XLWorkbook())
            using (var wb2 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");

                ws1.Cell("A1").FormulaA1 = "10*10";
                ws1.Cell("A2").FormulaA1 = "Original!A1 * 3";

                var ws2 = ws1.CopyTo(wb2, "Copy");

                Assert.AreEqual("Copy!A1 * 3", ws2.Cell("A2").FormulaA1);
            }
        }

        [Test, Ignore("Muted until #836 is fixed")]
        public void RenameWorksheetChangesAbsoluteReferencesInFormulae()
        {
            using (var wb1 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");

                ws1.Cell("A1").FormulaA1 = "10*10";
                ws1.Cell("A2").FormulaA1 = "Original!A1 * 3";

                ws1.Name = "Renamed";

                Assert.AreEqual("Renamed!A1 * 3", ws1.Cell("A2").FormulaA1);
            }
        }
    }
}
