using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using ClosedXML_Tests.Utils;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ClosedXML_Tests.Excel
{
    // Tests in this fixture test only the successful loading of existing Excel files,
    // i.e. we test that ClosedXML doesn't choke on a given input file
    // These tests DO NOT test that ClosedXML successfully recognises all the Excel parts or that it can successfully save those parts again.
    [TestFixture]
    public class LoadingTests
    {
        [Test]
        public void CanSuccessfullyLoadFiles()
        {
            var files = new List<string>()
            {
                @"Misc\TableWithCustomTheme.xlsx",
                @"Misc\EmptyTable.xlsx",
                @"Misc\LoadPivotTables.xlsx",
                @"Misc\LoadFileWithCustomSheetViews.xlsx",
                @"Misc\LoadSheetsWithCommas.xlsx",
                @"Misc\ExcelProducedWorkbookWithImages.xlsx",
                @"Misc\InvalidPrintTitles.xlsx",
                @"Misc\ExcelProducedWorkbookWithImages.xlsx",
                @"Misc\EmptyCellValue.xlsx",
                @"Misc\AllShapes.xlsx",
                @"Misc\TableHeadersWithLineBreaks.xlsx",
                @"Misc\TableWithNameNull.xlsx",
                @"Misc\DuplicateImageNames.xlsx",
                @"Misc\InvalidPrintArea.xlsx",
                @"Misc\Date1904System.xlsx",
                @"Misc\LoadImageWithoutTransform2D.xlsx",
                @"Misc\PivotTableWithTableSource.xlsx",
                @"Misc\TemplateWithTableSourcePivotTables.xlsx",
                @"Misc\PrintAreaRefersToExternalWorksheet.xlsx",
                @"Misc\NamedRangeWithInvalidCharacter.xlsx"
            };

            foreach (var file in files)
            {
                TestHelper.LoadFile(file);
            }
        }

        [Test]
        public void CanSuccessfullyLoadLOFiles()
        {
            // TODO: unpark all files
            var parkedForLater = new[]
            {
                "Misc.LO.xlsx.formats.xlsx",
                "Misc.LO.xlsx.pivot_table.shared-group-field.xlsx",
                "Misc.LO.xlsx.pivot_table.shared-nested-dategroup.xlsx",
                "Misc.LO.xlsx.pivottable_bool_field_filter.xlsx",
                "Misc.LO.xlsx.pivottable_date_field_filter.xlsx",
                "Misc.LO.xlsx.pivottable_double_field_filter.xlsx",
                "Misc.LO.xlsx.pivottable_duplicated_member_filter.xlsx",
                "Misc.LO.xlsx.pivottable_rowcolpage_field_filter.xlsx",
                "Misc.LO.xlsx.pivottable_string_field_filter.xlsx",
                "Misc.LO.xlsx.pivottable_tabular_mode.xlsx",
                "Misc.LO.xlsx.pivot_table_first_header_row.xlsx",
                "Misc.LO.xlsx.tdf100709.xlsx",
                "Misc.LO.xlsx.tdf89139_pivot_table.xlsx",
                "Misc.LO.xlsx.universal-content-strict.xlsx",
                "Misc.LO.xlsx.universal-content.xlsx",
                "Misc.LO.xlsx.xf_default_values.xlsx"
            };

            var files = TestHelper.ListResourceFiles(s => s.Contains(".LO.") && !parkedForLater.Any(i => s.Contains(i))).ToArray();

            foreach (var file in files)
            {
                TestHelper.LoadFile(file);
            }
        }

        [Test]
        public void CanLoadAndManipulateFileWithEmptyTable()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\EmptyTable.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                var table = ws.Tables.First();
                table.DataRange.InsertRowsBelow(5);
            }
        }

        [Test]
        public void CanLoadDate1904SystemCorrectly()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\Date1904System.xlsx")))
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook(stream))
                {
                    var ws = wb.Worksheets.First();
                    var c = ws.Cell("A2");
                    Assert.AreEqual(XLDataType.DateTime, c.DataType);
                    Assert.AreEqual(new DateTime(2017, 10, 27, 21, 0, 0), c.GetDateTime());
                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheets.First();
                    var c = ws.Cell("A2");
                    Assert.AreEqual(XLDataType.DateTime, c.DataType);
                    Assert.AreEqual(new DateTime(2017, 10, 27, 21, 0, 0), c.GetDateTime());
                    wb.SaveAs(ms);
                }
            }
        }

        [Test]
        public void CanLoadAndSaveFileWithMismatchingSheetIdAndRelId()
        {
            // This file's workbook.xml contains:
            // <x:sheet name="Data" sheetId="13" r:id="rId1" />
            // and the mismatch between the sheetId and r:id can create problems.
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\FileWithMismatchSheetIdAndRelId.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                using (var ms = new MemoryStream())
                {
                    wb.SaveAs(ms, true);
                }
            }
        }

        [Test]
        public void CanLoadBasicPivotTable()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\LoadPivotTables.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet("PivotTable1");
                var pt = ws.PivotTable("PivotTable1");
                Assert.AreEqual("PivotTable1", pt.Name);

                Assert.AreEqual(1, pt.RowLabels.Count());
                Assert.AreEqual("Name", pt.RowLabels.Single().SourceName);

                Assert.AreEqual(1, pt.ColumnLabels.Count());
                Assert.AreEqual("Month", pt.ColumnLabels.Single().SourceName);

                var pv = pt.Values.Single();
                Assert.AreEqual("Sum of NumberOfOrders", pv.CustomName);
                Assert.AreEqual("NumberOfOrders", pv.SourceName);
            }
        }

        [Test]
        public void CanLoadOrderedPivotTable()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\LoadPivotTables.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet("OrderedPivotTable");
                var pt = ws.PivotTable("OrderedPivotTable");

                Assert.AreEqual(XLPivotSortType.Ascending, pt.RowLabels.Single().SortType);
                Assert.AreEqual(XLPivotSortType.Descending, pt.ColumnLabels.Single().SortType);
            }
        }

        [Test]
        public void CanLoadPivotTableSubtotals()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\LoadPivotTables.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet("PivotTableSubtotals");
                var pt = ws.PivotTable("PivotTableSubtotals");

                var subtotals = pt.RowLabels.Get("Group").Subtotals.ToArray();
                Assert.AreEqual(3, subtotals.Length);
                Assert.AreEqual(XLSubtotalFunction.Average, subtotals[0]);
                Assert.AreEqual(XLSubtotalFunction.Count, subtotals[1]);
                Assert.AreEqual(XLSubtotalFunction.Sum, subtotals[2]);
            }
        }

        /// <summary>
        /// For non-English locales, the default style ("Normal" in English) can be
        /// another piece of text (e.g. ??????? in Russian).
        /// This test ensures that the default style is correctly detected and
        /// no style conflicts occur on save.
        /// </summary>
        [Test]
        public void CanSaveFileWithDefaultStyleNameNotInEnglish()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\FileWithDefaultStyleNameNotInEnglish.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                using (var ms = new MemoryStream())
                {
                    wb.SaveAs(ms, true);
                }
            }
        }

        /// <summary>
        /// As per https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.cellvalues(v=office.15).aspx
        /// the 'Date' DataType is available only in files saved with Microsoft Office
        /// In other files, the data type will be saved as numeric
        /// ClosedXML then deduces the data type by inspecting the number format string
        /// </summary>
        [Test]
        public void CanLoadLibreOfficeFileWithDates()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\LibreOfficeFileWithDates.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                foreach (var cell in ws.CellsUsed())
                {
                    Assert.AreEqual(XLDataType.DateTime, cell.DataType);
                }
            }
        }

        [Test]
        public void CanLoadFileWithImagesWithCorrectAnchorTypes()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                Assert.AreEqual(2, ws.Pictures.Count);
                Assert.AreEqual(XLPicturePlacement.FreeFloating, ws.Pictures.First().Placement);
                Assert.AreEqual(XLPicturePlacement.Move, ws.Pictures.Skip(1).First().Placement);

                var ws2 = wb.Worksheets.Skip(1).First();
                Assert.AreEqual(1, ws2.Pictures.Count);
                Assert.AreEqual(XLPicturePlacement.MoveAndSize, ws2.Pictures.First().Placement);
            }
        }

        [Test]
        public void CanLoadFileWithImagesWithCorrectImageType()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageFormats.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                Assert.AreEqual(1, ws.Pictures.Count);
                Assert.AreEqual(XLPictureFormat.Jpeg, ws.Pictures.First().Format);

                var ws2 = wb.Worksheets.Skip(1).First();
                Assert.AreEqual(1, ws2.Pictures.Count);
                Assert.AreEqual(XLPictureFormat.Png, ws2.Pictures.First().Format);
            }
        }

        [Test]
        public void CanLoadAndDeduceAnchorsFromExcelGeneratedFile()
        {
            // This file was produced by Excel. It contains 3 images, but the latter 2 were copied from the first.
            // There is actually only 1 embedded image if you inspect the file's internals.
            // Additionally, Excel saves all image anchors as TwoCellAnchor, but uses the EditAs attribute to distinguish the types
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\ExcelProducedWorkbookWithImages.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                Assert.AreEqual(3, ws.Pictures.Count);

                Assert.AreEqual(XLPicturePlacement.MoveAndSize, ws.Picture("Picture 1").Placement);
                Assert.AreEqual(XLPicturePlacement.Move, ws.Picture("Picture 2").Placement);
                Assert.AreEqual(XLPicturePlacement.FreeFloating, ws.Picture("Picture 3").Placement);

                using (var ms = new MemoryStream())
                    wb.SaveAs(ms, true);
            }
        }

        [Test]
        public void CanLoadFromTemplate()
        {
            using (var tf1 = new TemporaryFile())
            using (var tf2 = new TemporaryFile())
            {
                using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\AllShapes.xlsx")))
                using (var wb = new XLWorkbook(stream))
                {
                    // Save as temporary file
                    wb.SaveAs(tf1.Path);
                }

                var workbook = XLWorkbook.OpenFromTemplate(tf1.Path);
                Assert.True(workbook.Worksheets.Any());
                Assert.Throws<InvalidOperationException>(() => workbook.Save());

                workbook.SaveAs(tf2.Path);
            }
        }

        /// <summary>
        /// Excel escapes symbol ' in worksheet title so we have to process this correctly.
        /// </summary>
        [Test]
        public void CanOpenWorksheetWithEscapedApostrophe()
        {
            string title = "";
            TestDelegate openWorkbook = () =>
            {
                using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\EscapedApostrophe.xlsx")))
                using (var wb = new XLWorkbook(stream))
                {
                    var ws = wb.Worksheets.First();
                    title = ws.Name;
                }
            };

            Assert.DoesNotThrow(openWorkbook);
            Assert.AreEqual("L'E", title);
        }

        [Test]
        public void CanRoundTripSheetProtectionForObjects()
        {
            using (var book = new XLWorkbook())
            {
                var sheet = book.AddWorksheet("TestSheet");
                sheet.Protect()
                    .SetObjects(true)
                    .SetScenarios(true);

                using (var xlStream = new MemoryStream())
                {
                    book.SaveAs(xlStream);

                    using (var persistedBook = new XLWorkbook(xlStream))
                    {
                        var persistedSheet = persistedBook.Worksheets.Worksheet(1);

                        Assert.AreEqual(sheet.Protection.Objects, persistedSheet.Protection.Objects);
                    }
                }
            }
        }

        [Test]
        [TestCase("A1*10", "1230", 1230)]
        [TestCase("A1/10", "12.3", 12.3)]
        [TestCase("A1&\" cells\"", "123 cells", "123 cells")]
        [TestCase("A1&\"000\"", "123000", "123000")]
        [TestCase("ISNUMBER(A1)", "True", true)]
        [TestCase("ISBLANK(A1)", "False", false)]
        [TestCase("DATE(2018,1,28)", "43128", null)]
        public void LoadFormulaCachedValue(string formula, object expectedValueCached, object expectedCachedValue)
        {
            using (var ms = new MemoryStream())
            {
                using (XLWorkbook book1 = new XLWorkbook())
                {
                    var sheet = book1.AddWorksheet("sheet1");
                    sheet.Cell("A1").Value = 123;
                    sheet.Cell("A2").FormulaA1 = formula;
                    var options = new SaveOptions { EvaluateFormulasBeforeSaving = true };

                    book1.SaveAs(ms, options);
                }
                ms.Position = 0;

                using (XLWorkbook book2 = new XLWorkbook(ms))
                {
                    var ws = book2.Worksheet(1);
                    Assert.IsTrue(ws.Cell("A2").NeedsRecalculation);
                    Assert.AreEqual(expectedValueCached, ws.Cell("A2").ValueCached);

                    if (expectedCachedValue != null)
                        Assert.AreEqual(expectedCachedValue, ws.Cell("A2").CachedValue);
                }
            }
        }

        [Test]
        public void CanLoadWorksheetStyle()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\BaseColumnWidth.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet(1);

                Assert.AreEqual(8, ws.Style.Font.FontSize);
                Assert.AreEqual("Arial", ws.Style.Font.FontName);
                Assert.AreEqual(8, ws.Cell("A1").Style.Font.FontSize);
                Assert.AreEqual("Arial", ws.Cell("A1").Style.Font.FontName);
            }
        }

        [Test]
        public void CanCorrectLoadWorkbookCellWithStringDataType()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\CellWithStringDataType.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var cellToCheck = wb.Worksheet(1).Cell("B2");
                Assert.AreEqual(XLDataType.Text, cellToCheck.DataType);
                Assert.AreEqual("String with String Data type", cellToCheck.Value);
            }
        }

        [Test]
        public void CanLoadFileWithInvalidSelectedRanges()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\SelectedRanges\InvalidSelectedRange.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet(1);

                Assert.AreEqual(2, ws.SelectedRanges.Count);
                Assert.AreEqual("B2:B2", ws.SelectedRanges.First().RangeAddress.ToString());
                Assert.AreEqual("B2:C2", ws.SelectedRanges.Last().RangeAddress.ToString());
            }
        }
    }
}
