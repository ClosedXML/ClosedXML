using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using NUnit.Framework;
using System;
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
            Assert.That((end - start).TotalMilliseconds, Is.LessThan(500));
        }

        [Test]
        public void CopyConditionalFormatsCount()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Range("A1:C3").AddConditionalFormat().WhenContains("1").Fill.SetBackgroundColor(XLColor.Blue);
            ws.Range("A1:C3").Value = 1;
            IXLWorksheet ws2 = ws.CopyTo("Sheet2");
            Assert.That(ws2.ConditionalFormats.Count(), Is.EqualTo(1));
        }

        [Test]
        public void CopyColumnVisibility()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Columns(10, 20).Hide();
            ws.CopyTo("Sheet2");
            Assert.That(wb.Worksheet("Sheet2").Column(10).IsHidden, Is.True);
        }

        [Test]
        public void CopyRowVisibility()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Rows(2, 5).Hide();
            ws.CopyTo("Sheet2");
            Assert.That(wb.Worksheet("Sheet2").Row(4).IsHidden, Is.True);
        }

        [Test]
        public void DeletingSheets1()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet3");
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets.Add("Sheet1", 1);

            wb.Worksheet("Sheet3").Delete();

            Assert.Multiple(() =>
            {
                Assert.That(wb.Worksheet(1).Name, Is.EqualTo("Sheet1"));
                Assert.That(wb.Worksheet(2).Name, Is.EqualTo("Sheet2"));
                Assert.That(wb.Worksheets, Has.Count.EqualTo(2));
            });
        }

        [Test]
        public void InsertingSheets1()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet1");
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets.Add("Sheet3");

            Assert.Multiple(() =>
            {
                Assert.That(wb.Worksheet(1).Name, Is.EqualTo("Sheet1"));
                Assert.That(wb.Worksheet(2).Name, Is.EqualTo("Sheet2"));
                Assert.That(wb.Worksheet(3).Name, Is.EqualTo("Sheet3"));
            });
        }

        [Test]
        public void InsertingSheets2()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets.Add("Sheet1", 1);
            wb.Worksheets.Add("Sheet3");

            Assert.Multiple(() =>
            {
                Assert.That(wb.Worksheet(1).Name, Is.EqualTo("Sheet1"));
                Assert.That(wb.Worksheet(2).Name, Is.EqualTo("Sheet2"));
                Assert.That(wb.Worksheet(3).Name, Is.EqualTo("Sheet3"));
            });
        }

        [Test]
        public void InsertingSheets3()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet3");
            wb.Worksheets.Add("Sheet2", 1);
            wb.Worksheets.Add("Sheet1", 1);

            Assert.Multiple(() =>
            {
                Assert.That(wb.Worksheet(1).Name, Is.EqualTo("Sheet1"));
                Assert.That(wb.Worksheet(2).Name, Is.EqualTo("Sheet2"));
                Assert.That(wb.Worksheet(3).Name, Is.EqualTo("Sheet3"));
            });
        }

        [Test]
        public void InsertingSheets4()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add();

            Assert.That(ws1.Name, Is.EqualTo("Sheet1"));
            ws1.Name = "shEEt1";

            var ws2 = wb.Worksheets.Add();
            Assert.That(ws2.Name, Is.EqualTo("Sheet2"));

            wb.Worksheets.Add("SHEET4");

            Assert.That(wb.Worksheets.Add().Name, Is.EqualTo("Sheet5"));
            Assert.That(wb.Worksheets.Add().Name, Is.EqualTo("Sheet6"));

            wb.Worksheets.Add(1);

            Assert.That(wb.Worksheet(1).Name, Is.EqualTo("Sheet7"));
        }

        [Test]
        public void SheetIdIsNotReused()
        {
            using var wb = new XLWorkbook();
            var ws1 = (XLWorksheet)wb.AddWorksheet();
            var ws2 = (XLWorksheet)wb.AddWorksheet();
            var ws3 = (XLWorksheet)wb.AddWorksheet();

            Assert.Multiple(() =>
            {
                Assert.That(ws1.SheetId, Is.EqualTo(1));
                Assert.That(ws2.SheetId, Is.EqualTo(2));
                Assert.That(ws3.SheetId, Is.EqualTo(3));
            });

            ws3.Delete();
            var ws4 = (XLWorksheet)wb.AddWorksheet();
            Assert.That(ws4.SheetId, Is.EqualTo(4));
        }

        [Test]
        public void AddingDuplicateSheetNameThrowsException()
        {
            using var wb = new XLWorkbook();
            IXLWorksheet ws;
            ws = wb.AddWorksheet("Sheet1");

            Assert.Throws<ArgumentException>(() => wb.AddWorksheet("Sheet1"));

            //Sheet names are case insensitive
            Assert.Throws<ArgumentException>(() => wb.AddWorksheet("sheet1"));
        }

        [Test]
        public void MergedRanges()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("A1:B2").Merge();
            ws.Range("C1:D3").Merge();
            ws.Range("D2:E2").Merge();

            Assert.That(ws.MergedRanges, Has.Count.EqualTo(2));
            Assert.Multiple(() =>
            {
                Assert.That(ws.MergedRanges.First().RangeAddress.ToStringRelative(), Is.EqualTo("A1:B2"));
                Assert.That(ws.MergedRanges.Last().RangeAddress.ToStringRelative(), Is.EqualTo("D2:E2"));

                Assert.That(ws.Cell("A2").MergedRange().RangeAddress.ToStringRelative(), Is.EqualTo("A1:B2"));
                Assert.That(ws.Cell("D2").MergedRange().RangeAddress.ToStringRelative(), Is.EqualTo("D2:E2"));

                Assert.That(ws.Cell("Z10").MergedRange(), Is.Null);
            });
        }

        [Test]
        public void RowCountTime()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            DateTime start = DateTime.Now;
            ws.RowCount();
            DateTime end = DateTime.Now;
            Assert.That((end - start).TotalMilliseconds, Is.LessThan(500));
        }

        [Test]
        public void SheetsWithCommas()
        {
            using var wb = new XLWorkbook();
            var sourceSheetName = "Sheet1, Sheet3";
            var ws = wb.Worksheets.Add(sourceSheetName);
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Cell("B2").Value = 3;

            ws = wb.Worksheets.Add("Formula");
            ws.FirstCell().FormulaA1 = string.Format("=SUM('{0}'!A1:A2,'{0}'!B1:B2)", sourceSheetName);

            var value = ws.FirstCell().Value;
            Assert.That(value, Is.EqualTo(6));
        }

        [Test]
        public void CanRenameWorksheet()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet1");
            var ws2 = wb.AddWorksheet("Sheet2");

            ws1.Name = "New sheet name";
            Assert.That(ws1.Name, Is.EqualTo("New sheet name"));

            ws2.Name = "sheet2";
            Assert.That(ws2.Name, Is.EqualTo("sheet2"));

            Assert.Throws<ArgumentException>(() => ws1.Name = "SHEET2");
        }

        [Test]
        public void TryGetWorksheet()
        {
            using var wb = new XLWorkbook();
            wb.AddWorksheet("Sheet1");
            wb.AddWorksheet("Sheet2");

            IXLWorksheet ws;
            Assert.Multiple(() =>
            {
                Assert.That(wb.Worksheets.TryGetWorksheet("Sheet1", out ws), Is.True);
                Assert.That(wb.Worksheets.TryGetWorksheet("sheet1", out ws), Is.True);
                Assert.That(wb.Worksheets.TryGetWorksheet("sHEeT1", out ws), Is.True);
                Assert.That(wb.Worksheets.TryGetWorksheet("Sheeeet2", out ws), Is.False);

                Assert.That(wb.TryGetWorksheet("Sheet1", out ws), Is.True);
                Assert.That(wb.TryGetWorksheet("sheet1", out ws), Is.True);
                Assert.That(wb.TryGetWorksheet("sHEeT1", out ws), Is.True);
                Assert.That(wb.TryGetWorksheet("Sheeeet2", out ws), Is.False);
            });
        }

        [Test]
        public void HideWorksheet()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                wb.Worksheets.Add("VisibleSheet");
                wb.Worksheets.Add("HiddenSheet").Hide();
                wb.SaveAs(ms);
            }

            // unhide the hidden sheet
            using (var wb = new XLWorkbook(ms))
            {
                Assert.Multiple(() =>
                {
                    Assert.That(wb.Worksheet("VisibleSheet").Visibility, Is.EqualTo(XLWorksheetVisibility.Visible));
                    Assert.That(wb.Worksheet("HiddenSheet").Visibility, Is.EqualTo(XLWorksheetVisibility.Hidden));
                });

                var ws = wb.Worksheet("HiddenSheet");
                ws.Unhide().Name = "NoAlsoVisible";

                Assert.That(ws.Visibility, Is.EqualTo(XLWorksheetVisibility.Visible));

                wb.Save();
            }

            using (var wb = new XLWorkbook(ms))
            {
                Assert.Multiple(() =>
                {
                    Assert.That(wb.Worksheet("VisibleSheet").Visibility, Is.EqualTo(XLWorksheetVisibility.Visible));
                    Assert.That(wb.Worksheet("NoAlsoVisible").Visibility, Is.EqualTo(XLWorksheetVisibility.Visible));
                });
            }
        }

        [Test]
        public void CanCopySheetsWithAllAnchorTypes()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();
            ws.CopyTo("Copy1");

            var ws2 = wb.Worksheets.Skip(1).First();
            ws2.CopyTo("Copy2");

            var ws3 = wb.Worksheets.Skip(2).First();
            ws3.CopyTo("Copy3");

            var ws4 = wb.Worksheets.Skip(3).First();
            ws3.CopyTo("Copy4");
        }

        [Test]
        public void CannotCopyDeletedWorksheet()
        {
            using var wb = new XLWorkbook();
            wb.AddWorksheet("Sheet1");
            var ws = wb.AddWorksheet("Sheet2");

            ws.Delete();
            Assert.Throws<InvalidOperationException>(() => ws.CopyTo("Copy of Sheet2"));
        }

        [Test]
        public void WorksheetNameCannotStartWithApostrophe()
        {
            var title = "'StartsWithApostrophe";
            TestDelegate addWorksheet = () =>
            {
                using var wb = new XLWorkbook();
                wb.Worksheets.Add(title);
            };

            Assert.Throws(typeof(ArgumentException), addWorksheet);
        }

        [Test]
        public void WorksheetNameCannotEndWithApostrophe()
        {
            var title = "EndsWithApostrophe'";
            TestDelegate addWorksheet = () =>
            {
                using var wb = new XLWorkbook();
                wb.Worksheets.Add(title);
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
                using var ms = new MemoryStream();
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
            };

            Assert.DoesNotThrow(saveAndOpenWorkbook);
            Assert.That(savedTitle, Is.EqualTo(title));
        }

        [Test]
        public void CopyWorksheetPreservesContents()
        {
            using var wb1 = new XLWorkbook();
            using var wb2 = new XLWorkbook();
            var ws1 = wb1.Worksheets.Add("Original");

            ws1.Cell("A1").Value = "A1 value";
            ws1.Cell("A2").Value = 100;
            ws1.Cell("D4").Value = new DateTime(2018, 5, 1);

            var ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.Multiple(() =>
            {
                Assert.That(ws2.Cell("A1").Value, Is.EqualTo("A1 value"));
                Assert.That(ws2.Cell("A2").Value, Is.EqualTo(100));
                Assert.That(ws2.Cell("D4").Value.GetDateTime(), Is.EqualTo(new DateTime(2018, 5, 1)));
            });
        }

        [Test]
        public void CopyWorksheetPreservesFormulae()
        {
            using var wb1 = new XLWorkbook();
            using var wb2 = new XLWorkbook();
            var ws1 = wb1.Worksheets.Add("Original");

            ws1.Cell("A1").FormulaA1 = "10*10";
            ws1.Cell("A2").FormulaA1 = "A1 * 2";

            var ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.Multiple(() =>
            {
                Assert.That(ws2.Cell("A1").FormulaA1, Is.EqualTo("10*10"));
                Assert.That(ws2.Cell("A2").FormulaA1, Is.EqualTo("A1 * 2"));
            });
        }

        [Test]
        public void CopyWorksheetPreservesRowHeights()
        {
            using var wb1 = new XLWorkbook();
            var ws1 = wb1.Worksheets.Add("Original");
            using var wb2 = new XLWorkbook();
            ws1.RowHeight = 55;
            ws1.Row(2).Height = 0;
            ws1.Row(3).Height = 20;

            var ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.That(ws2.RowHeight, Is.EqualTo(ws1.RowHeight));
            for (int i = 1; i <= 3; i++)
            {
                Assert.That(ws2.Row(i).Height, Is.EqualTo(ws1.Row(i).Height));
            }
        }

        [Test]
        public void CopyWorksheetPreservesColumnWidths()
        {
            using var wb1 = new XLWorkbook();
            var ws1 = wb1.Worksheets.Add("Original");
            using var wb2 = new XLWorkbook();
            ws1.ColumnWidth = 160;
            ws1.Column(2).Width = 0;
            ws1.Column(3).Width = 240;

            var ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.That(ws2.ColumnWidth, Is.EqualTo(ws1.ColumnWidth));
            for (int i = 1; i <= 3; i++)
            {
                Assert.That(ws2.Column(i).Width, Is.EqualTo(ws1.Column(i).Width));
            }
        }

        [Test]
        public void CopyWorksheetPreservesMergedCells()
        {
            using var wb1 = new XLWorkbook();
            using var wb2 = new XLWorkbook();
            var ws1 = wb1.Worksheets.Add("Original");

            ws1.Range("A:A").Merge();
            ws1.Range("B1:C2").Merge();

            var ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.That(ws2.MergedRanges, Has.Count.EqualTo(ws1.MergedRanges.Count));
            for (int i = 0; i < ws1.MergedRanges.Count; i++)
            {
                Assert.That(ws2.MergedRanges.ElementAt(i).RangeAddress.ToString(),
                    Is.EqualTo(ws1.MergedRanges.ElementAt(i).RangeAddress.ToString()));
            }
        }

        [Test]
        public void Copy_sheet_across_workbooks_preserves_defined_names()
        {
            using var wb1 = new XLWorkbook();
            using var wb2 = new XLWorkbook();
            var ws1 = wb1.Worksheets.Add("Original");

            ws1.Range("A1:A2").AddToNamed("GLOBAL", XLScope.Workbook);
            ws1.Ranges("B1:B2,D1:D2").AddToNamed("LOCAL", XLScope.Worksheet);

            var ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.That(ws2.DefinedNames.Count(), Is.EqualTo(ws1.DefinedNames.Count()));
            for (int i = 0; i < ws1.DefinedNames.Count(); i++)
            {
                var nr1 = ws1.DefinedNames.ElementAt(i);
                var nr2 = ws2.DefinedNames.ElementAt(i);
                Assert.Multiple(() =>
                {
                    Assert.That(nr2.Ranges.ToString(), Is.EqualTo(nr1.Ranges.ToString()));
                    Assert.That(nr2.Scope, Is.EqualTo(nr1.Scope));
                    Assert.That(nr2.Name, Is.EqualTo(nr1.Name));
                    Assert.That(nr2.Visible, Is.EqualTo(nr1.Visible));
                    Assert.That(nr2.Comment, Is.EqualTo(nr1.Comment));
                });
            }
        }

        [Test]
        public void Copying_sheet_inside_workbook_makes_copies_of_sheet_scoped_defined_names()
        {
            using var wb1 = new XLWorkbook();
            var ws1 = wb1.Worksheets.Add("Original");

            ws1.Range("A1:A2").AddToNamed("GLOBAL", XLScope.Workbook);
            ws1.Ranges("B1:B2,D1:D2").AddToNamed("LOCAL", XLScope.Worksheet);

            var ws2 = ws1.CopyTo("Copy");

            Assert.That(ws2.DefinedNames.Count(), Is.EqualTo(ws1.DefinedNames.Count()));
            for (int i = 0; i < ws1.DefinedNames.Count(); i++)
            {
                var nr1 = ws1.DefinedNames.ElementAt(i);
                var nr2 = ws2.DefinedNames.ElementAt(i);
                Assert.Multiple(() =>
                {
                    Assert.That((int)nr2.Scope, Is.EqualTo((int)XLScope.Worksheet));
                    Assert.That(nr2.Ranges.ToString(), Is.EqualTo(nr1.Ranges.ToString()));
                    Assert.That(nr2.Name, Is.EqualTo(nr1.Name));
                    Assert.That(nr2.Visible, Is.EqualTo(nr1.Visible));
                    Assert.That(nr2.Comment, Is.EqualTo(nr1.Comment));
                });
            }
        }
        [Test]
        public void CopyWorksheetPreservesStyles()
        {
            using var ms = new MemoryStream();
            using var wb1 = new XLWorkbook();
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
                Assert.That((ws2.Style as XLStyle).Value, Is.EqualTo((ws1.Style as XLStyle).Value),
                    "Worksheet styles differ");
                var cellsUsed = ws1.Range(ws1.FirstCell(), ws1.LastCellUsed()).Cells();
                foreach (var cell in cellsUsed)
                {
                    var style1 = (cell.Style as XLStyle).Value;
                    var style2 = (ws2.Cell(cell.Address.ToString()).Style as XLStyle).Value;
                    Assert.That(style2, Is.EqualTo(style1), $"Cell {cell.Address} styles differ");
                }
            }
        }

        [Test]
        public void CopyWorksheetPreservesConditionalFormats()
        {
            using var wb1 = new XLWorkbook();
            using var wb2 = new XLWorkbook();
            var ws1 = wb1.Worksheets.Add("Original");

            ws1.Range("A:A").AddConditionalFormat()
                .WhenContains("0").Fill.SetBackgroundColor(XLColor.Red);
            var cf = ws1.Range("B1:C2").AddConditionalFormat();
            cf.Ranges.Add(ws1.Range("D4:D5"));
            cf.WhenEqualOrGreaterThan(100).Font.SetBold();

            var ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.That(ws2.ConditionalFormats.Count(), Is.EqualTo(ws1.ConditionalFormats.Count()));
            for (int i = 0; i < ws1.ConditionalFormats.Count(); i++)
            {
                var original = ws1.ConditionalFormats.ElementAt(i);
                var copy = ws2.ConditionalFormats.ElementAt(i);
                Assert.That(copy.Ranges, Has.Count.EqualTo(original.Ranges.Count));
                for (int j = 0; j < original.Ranges.Count; j++)
                {
                    Assert.That(copy.Ranges.ElementAt(j).RangeAddress.ToString(XLReferenceStyle.A1, false),
                        Is.EqualTo(original.Ranges.ElementAt(j).RangeAddress.ToString(XLReferenceStyle.A1, false)));
                }

                Assert.That((copy.Style as XLStyle).Value, Is.EqualTo((original.Style as XLStyle).Value));
                Assert.That(copy.Values.Single().Value.Value, Is.EqualTo(original.Values.Single().Value.Value));
            }
        }

        [Test]
        public void CopyWorksheetPreservesTables()
        {
            using var wb1 = new XLWorkbook();
            using var wb2 = new XLWorkbook();
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

            Assert.That(ws2.Tables.Count(), Is.EqualTo(ws1.Tables.Count()));
            for (int i = 0; i < ws1.Tables.Count(); i++)
            {
                var original = ws1.Tables.ElementAt(i);
                var copy = ws2.Tables.ElementAt(i);
                Assert.Multiple(() =>
                {
                    Assert.That(copy.RangeAddress.ToString(XLReferenceStyle.A1, false), Is.EqualTo(original.RangeAddress.ToString(XLReferenceStyle.A1, false)));
                    Assert.That(copy.Fields.Count(), Is.EqualTo(original.Fields.Count()));
                });
                for (int j = 0; j < original.Fields.Count(); j++)
                {
                    var originalField = original.Fields.ElementAt(j);
                    var copyField = copy.Fields.ElementAt(j);
                    Assert.Multiple(() =>
                    {
                        Assert.That(copyField.Name, Is.EqualTo(originalField.Name));
                        Assert.That(copyField.TotalsRowFormulaA1, Is.EqualTo(originalField.TotalsRowFormulaA1));
                        Assert.That(copyField.TotalsRowFunction, Is.EqualTo(originalField.TotalsRowFunction));
                    });
                }

                Assert.Multiple(() =>
                {
                    Assert.That(copy.Name, Is.EqualTo(original.Name));
                    Assert.That(copy.ShowAutoFilter, Is.EqualTo(original.ShowAutoFilter));
                    Assert.That(copy.ShowColumnStripes, Is.EqualTo(original.ShowColumnStripes));
                    Assert.That(copy.ShowHeaderRow, Is.EqualTo(original.ShowHeaderRow));
                    Assert.That(copy.ShowRowStripes, Is.EqualTo(original.ShowRowStripes));
                    Assert.That(copy.ShowTotalsRow, Is.EqualTo(original.ShowTotalsRow));
                    Assert.That((copy.Style as XLStyle).Value, Is.EqualTo((original.Style as XLStyle).Value));
                    Assert.That(copy.Theme, Is.EqualTo(original.Theme));
                });
            }
        }

        [Test]
        public void CopyWorksheetPreservesDataValidation()
        {
            using var wb1 = new XLWorkbook();
            using var wb2 = new XLWorkbook();
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

            Assert.That(ws2.DataValidations.Count(), Is.EqualTo(ws1.DataValidations.Count()));
            for (int i = 0; i < ws1.DataValidations.Count(); i++)
            {
                var original = ws1.DataValidations.ElementAt(i);
                var copy = ws2.DataValidations.ElementAt(i);

                var originalRanges = string.Join(",", original.Ranges.Select(r => r.RangeAddress.ToString()));
                var copyRanges = string.Join(",", original.Ranges.Select(r => r.RangeAddress.ToString()));

                Assert.Multiple(() =>
                {
                    Assert.That(copyRanges, Is.EqualTo(originalRanges));
                    Assert.That(copy.AllowedValues, Is.EqualTo(original.AllowedValues));
                    Assert.That(copy.Operator, Is.EqualTo(original.Operator));
                    Assert.That(copy.ErrorStyle, Is.EqualTo(original.ErrorStyle));
                    Assert.That(copy.ErrorTitle, Is.EqualTo(original.ErrorTitle));
                    Assert.That(copy.ErrorMessage, Is.EqualTo(original.ErrorMessage));
                });
            }
        }

        [Test]
        public void CopyWorksheetPreservesPictures()
        {
            using var ms = new MemoryStream();
            using var imageStream = Assembly.GetAssembly(typeof(ClosedXML.Examples.BasicTable))
                .GetManifestResourceStream("ClosedXML.Examples.Resources.SampleImage.jpg");
            using var wb1 = new XLWorkbook();
            {
                var ws1 = wb1.Worksheets.Add("Original");

                var picture = ws1.AddPicture(imageStream, "MyPicture")
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
                Assert.That(ws2.Pictures, Has.Count.EqualTo(ws1.Pictures.Count));

                for (var i = 0; i < ws1.Pictures.Count; i++)
                {
                    var original = ws1.Pictures.ElementAt(i);
                    var copy = ws2.Pictures.ElementAt(i);
                    Assert.Multiple(() =>
                    {
                        Assert.That(copy.Worksheet, Is.EqualTo(ws2));

                        Assert.That(copy.Format, Is.EqualTo(original.Format));
                        Assert.That(copy.Height, Is.EqualTo(original.Height));
                        Assert.That(copy.Id, Is.EqualTo(original.Id));
                        Assert.That(copy.Left, Is.EqualTo(original.Left));
                        Assert.That(copy.Name, Is.EqualTo(original.Name));
                        Assert.That(copy.Placement, Is.EqualTo(original.Placement));
                        Assert.That(copy.Top, Is.EqualTo(original.Top));
                        Assert.That(copy.TopLeftCell.Address.ToString(), Is.EqualTo(original.TopLeftCell.Address.ToString()));
                        Assert.That(copy.Width, Is.EqualTo(original.Width));
                        Assert.That(copy.ImageStream.ToArray(), Is.EqualTo(original.ImageStream.ToArray()), "Image streams differ");
                    });
                }
            }
        }

        [Test]
        public void CopyWorksheetPreservesPivotTables()
        {
            using var ms = new MemoryStream();
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx"));
            using var wb = new XLWorkbook(stream);
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
                Assert.That(ws2.PivotTables.Count(), Is.EqualTo(ws1.PivotTables.Count()));

                var comparer = new PivotTableComparer();

                for (int i = 0; i < ws1.PivotTables.Count(); i++)
                {
                    var original = ws1.PivotTables.ElementAt(i).CastTo<XLPivotTable>();
                    var copy = ws2.PivotTables.ElementAt(i).CastTo<XLPivotTable>();

                    Assert.Multiple(() =>
                    {
                        Assert.That(copy.Worksheet, Is.EqualTo(ws2));

                        Assert.That(comparer.Equals(original, copy), Is.True);
                    });
                }
            }
        }

        [Test]
        public void CopyWorksheetPreservesSelectedRanges()
        {
            using var wb1 = new XLWorkbook();
            using var wb2 = new XLWorkbook();
            var ws1 = wb1.Worksheets.Add("Original");

            ws1.SelectedRanges.RemoveAll();
            ws1.SelectedRanges.Add(ws1.Range("E12:H20"));
            ws1.SelectedRanges.Add(ws1.Range("B:B"));
            ws1.SelectedRanges.Add(ws1.Range("3:6"));

            var ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.That(ws2.SelectedRanges, Has.Count.EqualTo(ws1.SelectedRanges.Count));
            for (int i = 0; i < ws1.SelectedRanges.Count; i++)
            {
                Assert.That(ws2.SelectedRanges.ElementAt(i).RangeAddress.ToString(),
                    Is.EqualTo(ws1.SelectedRanges.ElementAt(i).RangeAddress.ToString()));
            }
        }

        [Test]
        public void CopyWorksheetPreservesPageSetup()
        {
            using var wb1 = new XLWorkbook();
            using var wb2 = new XLWorkbook();
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

            Assert.Multiple(() =>
            {
                Assert.That(ws2.PageSetup.FirstRowToRepeatAtTop, Is.EqualTo(ws1.PageSetup.FirstRowToRepeatAtTop));
                Assert.That(ws2.PageSetup.LastRowToRepeatAtTop, Is.EqualTo(ws1.PageSetup.LastRowToRepeatAtTop));
                Assert.That(ws2.PageSetup.FirstColumnToRepeatAtLeft, Is.EqualTo(ws1.PageSetup.FirstColumnToRepeatAtLeft));
                Assert.That(ws2.PageSetup.LastColumnToRepeatAtLeft, Is.EqualTo(ws1.PageSetup.LastColumnToRepeatAtLeft));
                Assert.That(ws2.PageSetup.PageOrientation, Is.EqualTo(ws1.PageSetup.PageOrientation));
                Assert.That(ws2.PageSetup.PagesWide, Is.EqualTo(ws1.PageSetup.PagesWide));
                Assert.That(ws2.PageSetup.PagesTall, Is.EqualTo(ws1.PageSetup.PagesTall));
                Assert.That(ws2.PageSetup.Scale, Is.EqualTo(ws1.PageSetup.Scale));
                Assert.That(ws2.PageSetup.HorizontalDpi, Is.EqualTo(ws1.PageSetup.HorizontalDpi));
                Assert.That(ws2.PageSetup.VerticalDpi, Is.EqualTo(ws1.PageSetup.VerticalDpi));
                Assert.That(ws2.PageSetup.FirstPageNumber, Is.EqualTo(ws1.PageSetup.FirstPageNumber));
                Assert.That(ws2.PageSetup.CenterHorizontally, Is.EqualTo(ws1.PageSetup.CenterHorizontally));
                Assert.That(ws2.PageSetup.CenterVertically, Is.EqualTo(ws1.PageSetup.CenterVertically));
                Assert.That(ws2.PageSetup.PaperSize, Is.EqualTo(ws1.PageSetup.PaperSize));
                Assert.That(ws2.PageSetup.Margins.Bottom, Is.EqualTo(ws1.PageSetup.Margins.Bottom));
                Assert.That(ws2.PageSetup.Margins.Top, Is.EqualTo(ws1.PageSetup.Margins.Top));
                Assert.That(ws2.PageSetup.Margins.Left, Is.EqualTo(ws1.PageSetup.Margins.Left));
                Assert.That(ws2.PageSetup.Margins.Right, Is.EqualTo(ws1.PageSetup.Margins.Right));
                Assert.That(ws2.PageSetup.Margins.Footer, Is.EqualTo(ws1.PageSetup.Margins.Footer));
                Assert.That(ws2.PageSetup.Margins.Header, Is.EqualTo(ws1.PageSetup.Margins.Header));
                Assert.That(ws2.PageSetup.ScaleHFWithDocument, Is.EqualTo(ws1.PageSetup.ScaleHFWithDocument));
                Assert.That(ws2.PageSetup.AlignHFWithMargins, Is.EqualTo(ws1.PageSetup.AlignHFWithMargins));
                Assert.That(ws2.PageSetup.ShowGridlines, Is.EqualTo(ws1.PageSetup.ShowGridlines));
                Assert.That(ws2.PageSetup.ShowRowAndColumnHeadings, Is.EqualTo(ws1.PageSetup.ShowRowAndColumnHeadings));
                Assert.That(ws2.PageSetup.BlackAndWhite, Is.EqualTo(ws1.PageSetup.BlackAndWhite));
                Assert.That(ws2.PageSetup.DraftQuality, Is.EqualTo(ws1.PageSetup.DraftQuality));
                Assert.That(ws2.PageSetup.PageOrder, Is.EqualTo(ws1.PageSetup.PageOrder));
                Assert.That(ws2.PageSetup.ShowComments, Is.EqualTo(ws1.PageSetup.ShowComments));
                Assert.That(ws2.PageSetup.PrintErrorValue, Is.EqualTo(ws1.PageSetup.PrintErrorValue));

                Assert.That(ws2.PageSetup.PrintAreas.Count(), Is.EqualTo(ws1.PageSetup.PrintAreas.Count()));

                Assert.That(ws2.PageSetup.Header.Left.GetText(XLHFOccurrence.AllPages), Is.EqualTo(ws1.PageSetup.Header.Left.GetText(XLHFOccurrence.AllPages)));
                Assert.That(ws2.PageSetup.Footer.Right.GetText(XLHFOccurrence.OddPages), Is.EqualTo(ws1.PageSetup.Footer.Right.GetText(XLHFOccurrence.OddPages)));
            });
        }

        [Test]
        public void CopyWorksheetPreservesSparklineGroups()
        {
            using var wb1 = new XLWorkbook();
            using var wb2 = new XLWorkbook();
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

            Assert.That(ws2.SparklineGroups.Count(), Is.EqualTo(1));
            var copy = ws2.SparklineGroups.Single();

            Assert.That(copy.Count(), Is.EqualTo(original.Count()));
            for (int i = 0; i < original.Count(); i++)
            {
                Assert.Multiple(() =>
                {
                    Assert.That(copy.ElementAt(i).Location.Worksheet, Is.SameAs(ws2));
                    Assert.That(copy.ElementAt(i).SourceData.Worksheet, Is.SameAs(ws2));
                    Assert.That(copy.ElementAt(i).Location.Address.ToString(), Is.EqualTo(original.ElementAt(i).Location.Address.ToString()));
                });
                Assert.That(copy.ElementAt(i).SourceData.RangeAddress.ToString(), Is.EqualTo(original.ElementAt(i).SourceData.RangeAddress.ToString()));
            }

            Assert.Multiple(() =>
            {
                Assert.That(copy.DateRange.RangeAddress.ToString(), Is.EqualTo(original.DateRange.RangeAddress.ToString()));
                Assert.That(copy.DateRange.Worksheet, Is.SameAs(ws2));
            });

            Assert.Multiple(() =>
            {
                Assert.That(copy.DisplayEmptyCellsAs, Is.EqualTo(original.DisplayEmptyCellsAs));
                Assert.That(copy.DisplayHidden, Is.EqualTo(original.DisplayHidden));
                Assert.That(copy.LineWeight, Is.EqualTo(original.LineWeight).Within(XLHelper.Epsilon));
                Assert.That(copy.ShowMarkers, Is.EqualTo(original.ShowMarkers));
                Assert.That(copy.Style, Is.EqualTo(original.Style));
            });
            Assert.Multiple(() =>
            {
                Assert.That(copy.Style, Is.Not.SameAs(original.Style));
                Assert.That(copy.Type, Is.EqualTo(original.Type));

                Assert.That(copy.HorizontalAxis.Color, Is.EqualTo(original.HorizontalAxis.Color));
                Assert.That(copy.HorizontalAxis.DateAxis, Is.EqualTo(original.HorizontalAxis.DateAxis));
                Assert.That(copy.HorizontalAxis.IsVisible, Is.EqualTo(original.HorizontalAxis.IsVisible));
                Assert.That(copy.HorizontalAxis.RightToLeft, Is.EqualTo(original.HorizontalAxis.RightToLeft));

                Assert.That(copy.VerticalAxis.ManualMax, Is.EqualTo(original.VerticalAxis.ManualMax));
                Assert.That(copy.VerticalAxis.ManualMin, Is.EqualTo(original.VerticalAxis.ManualMin));
                Assert.That(copy.VerticalAxis.MaxAxisType, Is.EqualTo(original.VerticalAxis.MaxAxisType));
                Assert.That(copy.VerticalAxis.MinAxisType, Is.EqualTo(original.VerticalAxis.MinAxisType));
            });
        }

        [Test, Ignore("Muted until #836 is fixed")]
        public void CopyWorksheetChangesAbsoluteReferencesInFormulae()
        {
            using var wb1 = new XLWorkbook();
            using var wb2 = new XLWorkbook();
            var ws1 = wb1.Worksheets.Add("Original");

            ws1.Cell("A1").FormulaA1 = "10*10";
            ws1.Cell("A2").FormulaA1 = "Original!A1 * 3";

            var ws2 = ws1.CopyTo(wb2, "Copy");

            Assert.That(ws2.Cell("A2").FormulaA1, Is.EqualTo("Copy!A1 * 3"));
        }

        [Test]
        public void Rename_sheets_changes_sheet_references_in_formulas()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Original");

            ws.Cell("A1").FormulaA1 = "10*10";
            ws.Cell("A2").FormulaA1 = "Original!A1 * 3";
            _ = ws.Cell("A2").Value;

            ws.Name = "Renamed";

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A2").FormulaA1, Is.EqualTo("Renamed!A1 * 3"));
                Assert.That(ws.Cell("A2").NeedsRecalculation, Is.True);
                Assert.That(ws.Cell("A2").Value, Is.EqualTo(300));
            });
        }

        [Test]
        public void RangesFromDeletedWorksheetContainREF()
        {
            using var wb1 = new XLWorkbook();
            wb1.Worksheets.Add("Sheet1");
            var ws2 = wb1.Worksheets.Add("Sheet2");
            var range = ws2.Range("A1:B2");

            ws2.Delete();

            Assert.That(range.RangeAddress.ToString(), Is.EqualTo("#REF!A1:B2"));
        }

        [Test]
        public void InvalidRowAndColumnIndices()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Row(-1));
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Row(XLHelper.MaxRowNumber + 1));

            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Column(-1));
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Column(XLHelper.MaxColumnNumber + 1));
        }

        [Test]
        public void InvalidSelectedRangeExcluded()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var range1 = ws.Range("B2:C2");
            var range2 = ws.Range("B4:C4");
            ws.SelectedRanges.Clear();

            ws.SelectedRanges.Add(range1);
            ws.SelectedRanges.Add(range2);

            ws.Row(4).Delete();

            Assert.Multiple(() =>
            {
                Assert.That(range2.RangeAddress.IsValid, Is.False);
                Assert.That(ws.SelectedRanges.Single(), Is.EqualTo(range1));
            });
        }

        [Test]
        public void InsertColumnsDoesNotIncreaseCellsCount()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").SetValue(1);
            ws.Cell("AAA50").SetValue(1);
            var originalCount = ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count();

            ws.Column(1).InsertColumnsBefore(1);

            Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count(), Is.EqualTo(originalCount));
        }

        [Test]
        public void InsertRowsDoesNotIncreaseCellsCount()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").SetValue(1);
            ws.Cell("AAA500").SetValue(1);
            var originalCount = ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count();

            ws.Row(1).InsertRowsAbove(1);

            Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count(), Is.EqualTo(originalCount));
        }

        [Test]
        public void InsertCellsBeforeDoesNotIncreaseCellsCount()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var a1 = ws.Cell("A1").SetValue(1);
            ws.Cell("AAA50").SetValue(1);
            var originalCount = ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count();

            a1.InsertCellsBefore(1);

            Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count(), Is.EqualTo(originalCount));
        }

        [Test]
        public void InsertCellsAboveDoesNotIncreaseCellsCount()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var a1 = ws.Cell("A1").SetValue(1);
            ws.Cell("AAA500").SetValue(1);
            var originalCount = ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count();

            a1.InsertCellsAbove(1);

            Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count(), Is.EqualTo(originalCount));
        }

        [Test]
        public void CellsShiftedTooFarRightArePurged()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var a1 = ws.Cell("A1").SetValue(1);
            ws.Cell(1, XLHelper.MaxColumnNumber).SetValue(1);
            ws.Cell(2, XLHelper.MaxColumnNumber).SetValue(1);

            a1.InsertCellsBefore(1);

            Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count(), Is.EqualTo(2));
            ws.Column(1).InsertColumnsBefore(1);
            Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count(), Is.EqualTo(1));
        }

        [Test]
        public void CellsShiftedTooFarDownArePurged()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var a1 = ws.Cell("A1").SetValue(1);
            ws.Cell(XLHelper.MaxRowNumber, 1).SetValue(1);
            ws.Cell(XLHelper.MaxRowNumber, 2).SetValue(1);

            a1.InsertCellsAbove(1);

            Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count(), Is.EqualTo(2));
            ws.Row(1).InsertRowsAbove(1);
            Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count(), Is.EqualTo(1));
        }

        [Test]
        public void MaxColumnUsedUpdatedWhenColumnDeleted()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("C1").SetValue(1);
            ws.Cell(1, XLHelper.MaxColumnNumber).SetValue(1);

            ws.Column(XLHelper.MaxColumnNumber).Delete();

            Assert.That(((XLWorksheet)ws).Internals.CellsCollection.MaxColumnUsed, Is.EqualTo(3));
        }

        [Test]
        public void MaxRowUsedUpdatedWhenRowDeleted()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A3").SetValue(1);
            ws.Cell(XLHelper.MaxRowNumber, 1).SetValue(1);

            ws.Row(XLHelper.MaxRowNumber).Delete();

            Assert.That(((XLWorksheet)ws).Internals.CellsCollection.MaxRowUsed, Is.EqualTo(3));
        }

        [Test]
        public void ChangeColumnStyleFirst()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("ColumnFirst");

            ws.Column(2).Style.Font.SetBold(true);
            ws.Row(2).Style.Font.SetItalic(true);

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("B2").Style.Font.Bold, Is.True);
                Assert.That(ws.Cell("B2").Style.Font.Italic, Is.True);
            });
        }

        [Test]
        public void ChangeRowStyleFirst()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("RowFirst");

            ws.Row(2).Style.Font.SetItalic(true);
            ws.Column(2).Style.Font.SetBold(true);

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("B2").Style.Font.Bold, Is.True);
                Assert.That(ws.Cell("B2").Style.Font.Italic, Is.True);
            });
        }

        [Test]
        public void SelectedTabIsActive_WhenInsertBefore()
        {
            using var ms = new MemoryStream();
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

                Assert.Multiple(() =>
                {
                    Assert.That(ws1.TabActive, Is.False);
                    Assert.That(ws1.TabSelected, Is.False);
                    Assert.That(ws2.TabActive, Is.True);
                    Assert.That(ws2.TabSelected, Is.True);
                });
            }
        }

        [TestCase("noactive_noselected.xlsx")]
        [TestCase("noactive_twoselected.xlsx")]
        [TestCase("noactive_negativeId.xlsx")]
        public void FirstSheetIsActive_WhenNotSpecified(string fileName)
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\NoActiveSheet\" + fileName));
            using var wb = new XLWorkbook(stream);
            Assert.Multiple(() =>
            {
                Assert.That(wb.Worksheets.First().TabActive, Is.True);
                Assert.That(wb.Worksheets.First().Visibility, Is.EqualTo(XLWorksheetVisibility.Visible));
            });
        }

        [TestCase(XLCellsUsedOptions.NormalFormats, 42)]
        [TestCase(XLCellsUsedOptions.Contents, 100)]
        public void FirstColumnUsed_ReturnsFirstColumnWithUsedCell(XLCellsUsedOptions options, int expectedColumn)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell(1, 42).Style.Fill.SetBackgroundColor(XLColor.Green);
            ws.Cell(1, 100).SetValue(5);

            var column = ws.FirstColumnUsed(options);
            Assert.That(column.ColumnNumber(), Is.EqualTo(expectedColumn));
        }

        [Test]
        public void RecalculateAllFormulas_recalculates_all_formulas_in_sheet_and_leaves_rest_dirty()
        {
            using var wb = new XLWorkbook();
            var sut = wb.AddWorksheet("sut");
            var other = wb.AddWorksheet("other");

            other.Cell("A1").Value = 7;
            other.Cell("A2").FormulaA1 = "A1+3";
            Assert.That(other.Cell("A2").Value, Is.EqualTo(10.0));

            // Change the supporting value, but without recalculation of dependent
            // formula, thus the value stays the same.
            other.Cell("A1").Value = 5;

            Assert.Multiple(() =>
            {
                Assert.That(other.Cell("A2").NeedsRecalculation, Is.True);
                Assert.That(other.Cell("A2").CachedValue, Is.EqualTo(10.0));
            });

            // Tested formula depends on a dirty formula from other sheet.
            sut.Cell("A1").FormulaA1 = "other!A2+5";
            sut.Cell("A2").FormulaA1 = "1+2";

            Assert.Multiple(() =>
            {
                Assert.That(sut.Cell("A1").CachedValue, Is.EqualTo(Blank.Value));
                Assert.That(sut.Cell("A2").CachedValue, Is.EqualTo(Blank.Value));
            });

            sut.RecalculateAllFormulas();

            Assert.Multiple(() =>
            {
                // Formulas in other sheets kept the value - not affected by recalculation of a sut sheet.
                Assert.That(other.Cell("A2").NeedsRecalculation, Is.True);
                Assert.That(other.Cell("A2").CachedValue, Is.EqualTo(10.0));
            });

            // Formulas in test sheet were recalculated - they are affected by recalculation of a sut sheet.
            Assert.That(sut.Cell("A1").NeedsRecalculation, Is.False);
            Assert.That(sut.Cell("A1").CachedValue, Is.EqualTo(15.0));

            Assert.That(sut.Cell("A2").NeedsRecalculation, Is.False);
            Assert.That(sut.Cell("A2").CachedValue, Is.EqualTo(3.0));
        }

        [Test]
        public void Cell_returns_cell_at_address_or_workbook_scoped_named_range()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            wb.DefinedNames.Add("test_range", ws.Range(2, 3, 5, 7)); // C2:G5

            var cellB4 = ws.Cell("B4");
            var firstCellOfRange = ws.Cell("test_range");

            Assert.Multiple(() =>
            {
                Assert.That(cellB4.Address.ToString(), Is.EqualTo("B4"));
                Assert.That(firstCellOfRange.Address.ToString(), Is.EqualTo("C2"));
            });
        }

        [Test]
        public void Cell_throws_exception_when_address_is_not_A1_address_or_workbook_scoped_range()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            Assert.Throws<ArgumentException>(() => _ = ws.Cell("XFF1"));
            Assert.Throws<ArgumentException>(() => _ = ws.Cell("nonexistent_range"));
        }

        [Test]
        public void Range_returns_range_from_a1_address_or_named_range()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            wb.DefinedNames.Add("book_range", ws.Range(2, 3, 5, 7)); // C2:G5
            ws.DefinedNames.Add("sheet_range", ws.Range(1, 2, 3, 4)); // B1:D3

            var singleCellRange = ws.Range("B4");
            var areaCellRange = ws.Range("B4:D7");
            var bookNamedRange = ws.Range("book_range");
            var sheetNamedRange = ws.Range("sheet_range");

            Assert.Multiple(() =>
            {
                Assert.That(singleCellRange.RangeAddress.ToString(), Is.EqualTo("B4:B4"));
                Assert.That(areaCellRange.RangeAddress.ToString(), Is.EqualTo("B4:D7"));
                Assert.That(bookNamedRange.RangeAddress.ToString(), Is.EqualTo("$C$2:$G$5"));
                Assert.That(sheetNamedRange.RangeAddress.ToString(), Is.EqualTo("$B$1:$D$3"));
            });
        }

        [Test]
        public void Range_throws_exception_when_address_is_not_A1_address_or_named_range()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            Assert.Throws<ArgumentException>(() => _ = ws.Range("DEAD1"));
            Assert.Throws<ArgumentException>(() => _ = ws.Range("DEAD4:BEEF10"));
            Assert.Throws<ArgumentException>(() => _ = ws.Range("nonexistent_range"));
        }
    }
}