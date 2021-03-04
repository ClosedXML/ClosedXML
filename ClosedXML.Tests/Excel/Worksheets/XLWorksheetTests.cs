using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using NUnit.Framework;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Tests
{
    [TestFixture]
    public class XLWorksheetTests
    {
        private readonly static char[] illegalWorksheetCharacters = "\u0000\u0003:\\/?*[]".ToCharArray();

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
            ws.Range("A1:C3").AddConditionalFormat().WhenContains("1").Fill.SetBackgroundColor(XLColor.Blue);
            ws.Range("A1:C3").Value = 1;
            IXLWorksheet ws2 = ws.CopyTo("Sheet2");
            Assert.AreEqual(1, ws2.ConditionalFormats.Count());
        }

        [Test]
        public void CopyColumnVisibility()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Columns(10, 20).Hide();
            ws.CopyTo("Sheet2");
            Assert.IsTrue(wb.Worksheet("Sheet2").Column(10).IsHidden);
        }

        [Test]
        public void CopyRowVisibility()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Rows(2, 5).Hide();
            ws.CopyTo("Sheet2");
            Assert.IsTrue(wb.Worksheet("Sheet2").Row(4).IsHidden);
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
        public void InsertingSheets4()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add();

            Assert.AreEqual("Sheet1", ws1.Name);
            ws1.Name = "shEEt1";

            var ws2 = wb.Worksheets.Add();
            Assert.AreEqual("Sheet2", ws2.Name);

            wb.Worksheets.Add("SHEET4");

            Assert.AreEqual("Sheet5", wb.Worksheets.Add().Name);
            Assert.AreEqual("Sheet6", wb.Worksheets.Add().Name);

            wb.Worksheets.Add(1);

            Assert.AreEqual("Sheet7", wb.Worksheet(1).Name);
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
        public void CannotCopyDeletedWorksheet()
        {
            using (var wb = new XLWorkbook())
            {
                wb.AddWorksheet("Sheet1");
                var ws = wb.AddWorksheet("Sheet2");

                ws.Delete();
                Assert.Throws<InvalidOperationException>(() => ws.CopyTo("Copy of Sheet2"));
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
        public void WorksheetNameCannotBeEmpty()
        {
            Assert.Throws<ArgumentException>(() => new XLWorkbook().AddWorksheet(" "));
        }

        [TestCaseSource(nameof(illegalWorksheetCharacters))]
        public void WorksheetNameCannotContainIllegalCharacters(char c)
        {
            var proposedName = $"Sheet{c}Name";
            Assert.Throws<ArgumentException>(() => new XLWorkbook().AddWorksheet(proposedName));
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
        public void CopyWorksheetPreservesRowHeights()
        {
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
                }
            }
        }

        [Test]
        public void CopyWorksheetPreservesColumnWidths()
        {
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
                var cf = ws1.Range("B1:C2").AddConditionalFormat();
                cf.Ranges.Add(ws1.Range("D4:D5"));
                cf.WhenEqualOrGreaterThan(100).Font.SetBold();

                var ws2 = ws1.CopyTo(wb2, "Copy");

                Assert.AreEqual(ws1.ConditionalFormats.Count(), ws2.ConditionalFormats.Count());
                for (int i = 0; i < ws1.ConditionalFormats.Count(); i++)
                {
                    var original = ws1.ConditionalFormats.ElementAt(i);
                    var copy = ws2.ConditionalFormats.ElementAt(i);
                    Assert.AreEqual(original.Ranges.Count, copy.Ranges.Count);
                    for (int j = 0; j < original.Ranges.Count; j++)
                    {
                        Assert.AreEqual(original.Ranges.ElementAt(j).RangeAddress.ToString(XLReferenceStyle.A1, false),
                            copy.Ranges.ElementAt(j).RangeAddress.ToString(XLReferenceStyle.A1, false));
                    }

                    Assert.AreEqual((original.Style as XLStyle).Value, (copy.Style as XLStyle).Value);
                    Assert.AreEqual(original.Values.Single().Value.Value, copy.Values.Single().Value.Value);
                }
            }
        }

        [Test]
        public void CopyWorksheetPreservesTables()
        {
            using (var wb1 = new XLWorkbook())
            using (var wb2 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");

                ws1.Cell("A2").Value = "Name";
                ws1.Cell("B2").Value = "Count";
                ws1.Cell("A3").Value = "John Smith";
                ws1.Cell("B3").Value = 50;
                ws1.Cell("A4").Value = "Ivan Ivanov";
                ws1.Cell("B4").Value = 40;
                var table1 = ws1.Range("A2:B4").CreateTable("Test table 1");
                table1
                    .SetShowAutoFilter(true)
                    .SetShowTotalsRow(true)
                    .SetEmphasizeFirstColumn(true)
                    .SetShowColumnStripes(true)
                    .SetShowRowStripes(true);
                table1.Theme = XLTableTheme.TableStyleDark8;
                table1.Field(1).TotalsRowFunction = XLTotalsRowFunction.Sum;

                var ws2 = ws1.CopyTo(wb2, "Copy");

                Assert.AreEqual(ws1.Tables.Count(), ws2.Tables.Count());
                for (int i = 0; i < ws1.Tables.Count(); i++)
                {
                    var original = ws1.Tables.ElementAt(i);
                    var copy = ws2.Tables.ElementAt(i);
                    Assert.AreEqual(original.RangeAddress.ToString(XLReferenceStyle.A1, false), copy.RangeAddress.ToString(XLReferenceStyle.A1, false));
                    Assert.AreEqual(original.Fields.Count(), copy.Fields.Count());
                    for (int j = 0; j < original.Fields.Count(); j++)
                    {
                        var originalField = original.Fields.ElementAt(j);
                        var copyField = copy.Fields.ElementAt(j);
                        Assert.AreEqual(originalField.Name, copyField.Name);
                        Assert.AreEqual(originalField.TotalsRowFormulaA1, copyField.TotalsRowFormulaA1);
                        Assert.AreEqual(originalField.TotalsRowFunction, copyField.TotalsRowFunction);
                    }

                    Assert.AreEqual(original.Name, copy.Name);
                    Assert.AreEqual(original.ShowAutoFilter, copy.ShowAutoFilter);
                    Assert.AreEqual(original.ShowColumnStripes, copy.ShowColumnStripes);
                    Assert.AreEqual(original.ShowHeaderRow, copy.ShowHeaderRow);
                    Assert.AreEqual(original.ShowRowStripes, copy.ShowRowStripes);
                    Assert.AreEqual(original.ShowTotalsRow, copy.ShowTotalsRow);
                    Assert.AreEqual((original.Style as XLStyle).Value, (copy.Style as XLStyle).Value);
                    Assert.AreEqual(original.Theme, copy.Theme);
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

                var dv1 = ws1.Range("A:A").CreateDataValidation();
                dv1.WholeNumber.EqualTo(2);
                dv1.ErrorStyle = XLErrorStyle.Warning;
                dv1.ErrorTitle = "Number out of range";
                dv1.ErrorMessage = "This cell only allows the number 2.";

                var dv2 = ws1.Ranges("B2:C3,D4:E5").CreateDataValidation();
                dv2.Decimal.GreaterThan(5);
                dv2.ErrorStyle = XLErrorStyle.Stop;
                dv2.ErrorTitle = "Decimal number out of range";
                dv2.ErrorMessage = "This cell only allows decimals greater than 5.";

                var dv3 = ws1.Cell("D1").CreateDataValidation();
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

                    var originalRanges = string.Join(",", original.Ranges.Select(r => r.RangeAddress.ToString()));
                    var copyRanges = string.Join(",", original.Ranges.Select(r => r.RangeAddress.ToString()));

                    Assert.AreEqual(originalRanges, copyRanges);
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
            using (var resourceStream = Assembly.GetAssembly(typeof(ClosedXML.Examples.BasicTable))
                .GetManifestResourceStream("ClosedXML.Examples.Resources.SampleImage.jpg"))
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
                    Assert.AreEqual(original.TopLeftCell.Address.ToString(), copy.TopLeftCell.Address.ToString());
                    Assert.AreEqual(original.Width, copy.Width);
                    Assert.AreEqual(original.ImageStream.ToArray(), copy.ImageStream.ToArray(), "Image streams differ");
                }
            }
        }

        [Test]
        public void CopyWorksheetPreservesPivotTables()
        {
            using (var ms = new MemoryStream())
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws1 = wb.Worksheet("pvt1");
                var copyOfws1 = ws1.CopyTo("CopyOfPvt1");

                AssertPivotTablesAreEqual(ws1, copyOfws1);

                using (var wb2 = new XLWorkbook())
                {
                    // We need to  copy the source too. Cross workbook references don't work yet.
                    wb.Worksheet("PastrySalesData").CopyTo(wb2);
                    var ws2 = ws1.CopyTo(wb2, "Copy");
                    AssertPivotTablesAreEqual(ws1, ws2);
                    wb2.SaveAs(ms);
                }

                using (var wb2 = new XLWorkbook(ms))
                {
                    var ws2 = wb2.Worksheet("Copy");
                    AssertPivotTablesAreEqual(ws1, ws2);
                }
            }

            void AssertPivotTablesAreEqual(IXLWorksheet ws1, IXLWorksheet ws2)
            {
                Assert.AreEqual(ws1.PivotTables.Count(), ws2.PivotTables.Count());

                var comparer = new PivotTableComparer();

                for (int i = 0; i < ws1.PivotTables.Count(); i++)
                {
                    var original = ws1.PivotTables.ElementAt(i).CastTo<XLPivotTable>();
                    var copy = ws2.PivotTables.ElementAt(i).CastTo<XLPivotTable>();

                    Assert.AreEqual(ws2, copy.Worksheet);
                    Assert.AreNotEqual(original.Guid, copy.Guid);

                    Assert.IsTrue(comparer.Equals(original, copy));
                }
            }
        }

        [Test]
        public void CopyWorksheetPreservesSelectedRanges()
        {
            using (var wb1 = new XLWorkbook())
            using (var wb2 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");

                ws1.SelectedRanges.RemoveAll();
                ws1.SelectedRanges.Add(ws1.Range("E12:H20"));
                ws1.SelectedRanges.Add(ws1.Range("B:B"));
                ws1.SelectedRanges.Add(ws1.Range("3:6"));

                var ws2 = ws1.CopyTo(wb2, "Copy");

                Assert.AreEqual(ws1.SelectedRanges.Count, ws2.SelectedRanges.Count);
                for (int i = 0; i < ws1.SelectedRanges.Count; i++)
                {
                    Assert.AreEqual(ws1.SelectedRanges.ElementAt(i).RangeAddress.ToString(),
                                    ws2.SelectedRanges.ElementAt(i).RangeAddress.ToString());
                }
            }
        }

        [Test]
        public void CopyWorksheetPreservesPageSetup()
        {
            using (var wb1 = new XLWorkbook())
            using (var wb2 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");

                ws1.PageSetup.AddHorizontalPageBreak(15);
                ws1.PageSetup.AddVerticalPageBreak(5);
                ws1.PageSetup
                    .SetBlackAndWhite()
                    .SetCenterHorizontally()
                    .SetCenterVertically()
                    .SetFirstPageNumber(200)
                    .SetPageOrientation(XLPageOrientation.Landscape)
                    .SetPaperSize(XLPaperSize.A5Paper)
                    .SetScale(89)
                    .SetShowGridlines()
                    .SetHorizontalDpi(200)
                    .SetVerticalDpi(300)
                    .SetPagesTall(5)
                    .SetPagesWide(2)
                    .SetColumnsToRepeatAtLeft(1, 3);
                ws1.PageSetup.PrintAreas.Clear();
                ws1.PageSetup.PrintAreas.Add("A1:Z200");
                ws1.PageSetup.Margins.SetBottom(5).SetTop(6).SetLeft(7).SetRight(8).SetFooter(9).SetHeader(10);
                ws1.PageSetup.Header.Left.AddText(XLHFPredefinedText.FullPath, XLHFOccurrence.AllPages);
                ws1.PageSetup.Footer.Right.AddText(XLHFPredefinedText.PageNumber, XLHFOccurrence.OddPages);

                var ws2 = ws1.CopyTo(wb2, "Copy");

                Assert.AreEqual(ws1.PageSetup.FirstRowToRepeatAtTop, ws2.PageSetup.FirstRowToRepeatAtTop);
                Assert.AreEqual(ws1.PageSetup.LastRowToRepeatAtTop, ws2.PageSetup.LastRowToRepeatAtTop);
                Assert.AreEqual(ws1.PageSetup.FirstColumnToRepeatAtLeft, ws2.PageSetup.FirstColumnToRepeatAtLeft);
                Assert.AreEqual(ws1.PageSetup.LastColumnToRepeatAtLeft, ws2.PageSetup.LastColumnToRepeatAtLeft);
                Assert.AreEqual(ws1.PageSetup.PageOrientation, ws2.PageSetup.PageOrientation);
                Assert.AreEqual(ws1.PageSetup.PagesWide, ws2.PageSetup.PagesWide);
                Assert.AreEqual(ws1.PageSetup.PagesTall, ws2.PageSetup.PagesTall);
                Assert.AreEqual(ws1.PageSetup.Scale, ws2.PageSetup.Scale);
                Assert.AreEqual(ws1.PageSetup.HorizontalDpi, ws2.PageSetup.HorizontalDpi);
                Assert.AreEqual(ws1.PageSetup.VerticalDpi, ws2.PageSetup.VerticalDpi);
                Assert.AreEqual(ws1.PageSetup.FirstPageNumber, ws2.PageSetup.FirstPageNumber);
                Assert.AreEqual(ws1.PageSetup.CenterHorizontally, ws2.PageSetup.CenterHorizontally);
                Assert.AreEqual(ws1.PageSetup.CenterVertically, ws2.PageSetup.CenterVertically);
                Assert.AreEqual(ws1.PageSetup.PaperSize, ws2.PageSetup.PaperSize);
                Assert.AreEqual(ws1.PageSetup.Margins.Bottom, ws2.PageSetup.Margins.Bottom);
                Assert.AreEqual(ws1.PageSetup.Margins.Top, ws2.PageSetup.Margins.Top);
                Assert.AreEqual(ws1.PageSetup.Margins.Left, ws2.PageSetup.Margins.Left);
                Assert.AreEqual(ws1.PageSetup.Margins.Right, ws2.PageSetup.Margins.Right);
                Assert.AreEqual(ws1.PageSetup.Margins.Footer, ws2.PageSetup.Margins.Footer);
                Assert.AreEqual(ws1.PageSetup.Margins.Header, ws2.PageSetup.Margins.Header);
                Assert.AreEqual(ws1.PageSetup.ScaleHFWithDocument, ws2.PageSetup.ScaleHFWithDocument);
                Assert.AreEqual(ws1.PageSetup.AlignHFWithMargins, ws2.PageSetup.AlignHFWithMargins);
                Assert.AreEqual(ws1.PageSetup.ShowGridlines, ws2.PageSetup.ShowGridlines);
                Assert.AreEqual(ws1.PageSetup.ShowRowAndColumnHeadings, ws2.PageSetup.ShowRowAndColumnHeadings);
                Assert.AreEqual(ws1.PageSetup.BlackAndWhite, ws2.PageSetup.BlackAndWhite);
                Assert.AreEqual(ws1.PageSetup.DraftQuality, ws2.PageSetup.DraftQuality);
                Assert.AreEqual(ws1.PageSetup.PageOrder, ws2.PageSetup.PageOrder);
                Assert.AreEqual(ws1.PageSetup.ShowComments, ws2.PageSetup.ShowComments);
                Assert.AreEqual(ws1.PageSetup.PrintErrorValue, ws2.PageSetup.PrintErrorValue);

                Assert.AreEqual(ws1.PageSetup.PrintAreas.Count(), ws2.PageSetup.PrintAreas.Count());

                Assert.AreEqual(ws1.PageSetup.Header.Left.GetText(XLHFOccurrence.AllPages), ws2.PageSetup.Header.Left.GetText(XLHFOccurrence.AllPages));
                Assert.AreEqual(ws1.PageSetup.Footer.Right.GetText(XLHFOccurrence.OddPages), ws2.PageSetup.Footer.Right.GetText(XLHFOccurrence.OddPages));
            }
        }

        [Test]
        public void CopyWorksheetPreservesSparklineGroups()
        {
            using (var wb1 = new XLWorkbook())
            using (var wb2 = new XLWorkbook())
            {
                var ws1 = wb1.Worksheets.Add("Original");
                var original = ws1.SparklineGroups.Add("A1:A10", "D1:Z10")
                    .SetDateRange(ws1.Range("D11:Z11"))
                    .SetDisplayEmptyCellsAs(XLDisplayBlanksAsValues.Zero)
                    .SetDisplayHidden(true)
                    .SetLineWeight(1.5)
                    .SetShowMarkers(XLSparklineMarkers.All)
                    .SetStyle(XLSparklineTheme.Colorful3)
                    .SetType(XLSparklineType.Column);

                original.HorizontalAxis
                    .SetColor(XLColor.Blue)
                    .SetRightToLeft(true)
                    .SetVisible(true);

                original.VerticalAxis
                    .SetManualMin(-100.0)
                    .SetManualMax(100.0);

                var ws2 = ws1.CopyTo(wb2, "Copy");

                Assert.AreEqual(1, ws2.SparklineGroups.Count());
                var copy = ws2.SparklineGroups.Single();

                Assert.AreEqual(original.Count(), copy.Count());
                for (int i = 0; i < original.Count(); i++)
                {
                    Assert.AreSame(ws2, copy.ElementAt(i).Location.Worksheet);
                    Assert.AreSame(ws2, copy.ElementAt(i).SourceData.Worksheet);
                    Assert.AreEqual(original.ElementAt(i).Location.Address.ToString(), copy.ElementAt(i).Location.Address.ToString());
                    Assert.AreEqual(original.ElementAt(i).SourceData.RangeAddress.ToString(), copy.ElementAt(i).SourceData.RangeAddress.ToString());
                }

                Assert.AreEqual(original.DateRange.RangeAddress.ToString(), copy.DateRange.RangeAddress.ToString());
                Assert.AreSame(ws2, copy.DateRange.Worksheet);

                Assert.AreEqual(original.DisplayEmptyCellsAs, copy.DisplayEmptyCellsAs);
                Assert.AreEqual(original.DisplayHidden, copy.DisplayHidden);
                Assert.AreEqual(original.LineWeight, copy.LineWeight, XLHelper.Epsilon);
                Assert.AreEqual(original.ShowMarkers, copy.ShowMarkers);
                Assert.AreEqual(original.Style, copy.Style);
                Assert.AreNotSame(original.Style, copy.Style);
                Assert.AreEqual(original.Type, copy.Type);

                Assert.AreEqual(original.HorizontalAxis.Color, copy.HorizontalAxis.Color);
                Assert.AreEqual(original.HorizontalAxis.DateAxis, copy.HorizontalAxis.DateAxis);
                Assert.AreEqual(original.HorizontalAxis.IsVisible, copy.HorizontalAxis.IsVisible);
                Assert.AreEqual(original.HorizontalAxis.RightToLeft, copy.HorizontalAxis.RightToLeft);

                Assert.AreEqual(original.VerticalAxis.ManualMax, copy.VerticalAxis.ManualMax);
                Assert.AreEqual(original.VerticalAxis.ManualMin, copy.VerticalAxis.ManualMin);
                Assert.AreEqual(original.VerticalAxis.MaxAxisType, copy.VerticalAxis.MaxAxisType);
                Assert.AreEqual(original.VerticalAxis.MinAxisType, copy.VerticalAxis.MinAxisType);
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

        [Test]
        public void RangesFromDeletedWorksheetContainREF()
        {
            using (var wb1 = new XLWorkbook())
            {
                wb1.Worksheets.Add("Sheet1");
                var ws2 = wb1.Worksheets.Add("Sheet2");
                var range = ws2.Range("A1:B2");

                ws2.Delete();

                Assert.AreEqual("#REF!A1:B2", range.RangeAddress.ToString());
            }
        }

        [Test]
        public void InvalidRowAndColumnIndices()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                Assert.Throws<ArgumentOutOfRangeException>(() => ws.Row(-1));
                Assert.Throws<ArgumentOutOfRangeException>(() => ws.Row(XLHelper.MaxRowNumber + 1));

                Assert.Throws<ArgumentOutOfRangeException>(() => ws.Column(-1));
                Assert.Throws<ArgumentOutOfRangeException>(() => ws.Column(XLHelper.MaxColumnNumber + 1));
            }
        }

        [Test]
        public void InvalidSelectedRangeExcluded()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                var range1 = ws.Range("B2:C2");
                var range2 = ws.Range("B4:C4");
                ws.SelectedRanges.Clear();

                ws.SelectedRanges.Add(range1);
                ws.SelectedRanges.Add(range2);

                ws.Row(4).Delete();

                Assert.IsFalse(range2.RangeAddress.IsValid);
                Assert.AreEqual(range1, ws.SelectedRanges.Single());
            }
        }

        [Test]
        public void InsertColumnsDoesNotIncreaseCellsCount()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var cell1 = ws.Cell("A1");
                var cell2 = ws.Cell("AAA50");
                var originalCount = (ws as XLWorksheet).Internals.CellsCollection.Count;

                ws.Column(1).InsertColumnsBefore(1);

                Assert.AreEqual(originalCount, (ws as XLWorksheet).Internals.CellsCollection.Count);
            }
        }

        [Test]
        public void InsertRowsDoesNotIncreaseCellsCount()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var cell1 = ws.Cell("A1");
                var cell2 = ws.Cell("AAA500");
                var originalCount = (ws as XLWorksheet).Internals.CellsCollection.Count;

                ws.Row(1).InsertRowsAbove(1);

                Assert.AreEqual(originalCount, (ws as XLWorksheet).Internals.CellsCollection.Count);
            }
        }

        [Test]
        public void InsertCellsBeforeDoesNotIncreaseCellsCount()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var cell1 = ws.Cell("A1");
                var cell2 = ws.Cell("AAA50");
                var originalCount = (ws as XLWorksheet).Internals.CellsCollection.Count;

                cell1.InsertCellsBefore(1);

                Assert.AreEqual(originalCount, (ws as XLWorksheet).Internals.CellsCollection.Count);
            }
        }

        [Test]
        public void InsertCellsAboveDoesNotIncreaseCellsCount()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var cell1 = ws.Cell("A1");
                var cell2 = ws.Cell("AAA500");
                var originalCount = (ws as XLWorksheet).Internals.CellsCollection.Count;

                cell1.InsertCellsAbove(1);

                Assert.AreEqual(originalCount, (ws as XLWorksheet).Internals.CellsCollection.Count);
            }
        }

        [Test]
        public void CellsShiftedTooFarRightArePurged()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var cell1 = ws.Cell("A1");
                var cell2 = ws.Cell(1, XLHelper.MaxColumnNumber);
                var cell3 = ws.Cell(2, XLHelper.MaxColumnNumber);

                cell1.InsertCellsBefore(1);
                Assert.AreEqual(2, (ws as XLWorksheet).Internals.CellsCollection.Count);
                ws.Column(1).InsertColumnsBefore(1);
                Assert.AreEqual(1, (ws as XLWorksheet).Internals.CellsCollection.Count);
            }
        }

        [Test]
        public void CellsShiftedTooFarDownArePurged()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var cell1 = ws.Cell("A1");
                var cell2 = ws.Cell(XLHelper.MaxRowNumber, 1);
                var cell3 = ws.Cell(XLHelper.MaxRowNumber, 2);

                cell1.InsertCellsAbove(1);
                Assert.AreEqual(2, (ws as XLWorksheet).Internals.CellsCollection.Count);
                ws.Row(1).InsertRowsAbove(1);
                Assert.AreEqual(1, (ws as XLWorksheet).Internals.CellsCollection.Count);
            }
        }

        [Test]
        public void MaxColumnUsedUpdatedWhenColumnDeleted()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var cell1 = ws.Cell("C1");
                var cell2 = ws.Cell(1, XLHelper.MaxColumnNumber);

                ws.Column(XLHelper.MaxColumnNumber).Delete();

                Assert.AreEqual(3, (ws as XLWorksheet).Internals.CellsCollection.MaxColumnUsed);
            }
        }

        [Test]
        public void MaxRowUsedUpdatedWhenRowDeleted()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var cell1 = ws.Cell("A3");
                var cell2 = ws.Cell(XLHelper.MaxRowNumber, 1);

                ws.Row(XLHelper.MaxRowNumber).Delete();

                Assert.AreEqual(3, (ws as XLWorksheet).Internals.CellsCollection.MaxRowUsed);
            }
        }

        [Test]
        public void ChangeColumnStyleFirst()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("ColumnFirst");

                ws.Column(2).Style.Font.SetBold(true);
                ws.Row(2).Style.Font.SetItalic(true);

                Assert.IsTrue(ws.Cell("B2").Style.Font.Bold);
                Assert.IsTrue(ws.Cell("B2").Style.Font.Italic);
            }
        }

        [Test]
        public void ChangeRowStyleFirst()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("RowFirst");

                ws.Row(2).Style.Font.SetItalic(true);
                ws.Column(2).Style.Font.SetBold(true);

                Assert.IsTrue(ws.Cell("B2").Style.Font.Bold);
                Assert.IsTrue(ws.Cell("B2").Style.Font.Italic);
            }
        }

        [Test]
        public void SelectedTabIsActive_WhenInsertBefore()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws1 = wb.AddWorksheet();
                    ws1.TabSelected = true;
                    var ws2 = wb.Worksheets.Add(1);
                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var ws1 = wb.Worksheets.First();
                    var ws2 = wb.Worksheets.Last();

                    Assert.IsFalse(ws1.TabActive);
                    Assert.IsFalse(ws1.TabSelected);
                    Assert.IsTrue(ws2.TabActive);
                    Assert.IsTrue(ws2.TabSelected);
                }
            }
        }

        [TestCase("noactive_noselected.xlsx")]
        [TestCase("noactive_twoselected.xlsx")]
        public void FirstSheetIsActive_WhenNotSpecified(string fileName)
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\NoActiveSheet\" + fileName)))
            using (var wb = new XLWorkbook(stream))
            {
                Assert.IsTrue(wb.Worksheets.First().TabActive);
                Assert.AreEqual(XLWorksheetVisibility.Visible, wb.Worksheets.First().Visibility);
            }
        }
    }
}
