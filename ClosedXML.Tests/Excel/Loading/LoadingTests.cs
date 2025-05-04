using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using ClosedXML.Tests.Utils;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel
{
    // Tests in this fixture test only the successful loading of existing Excel files,
    // i.e. we test that ClosedXML doesn't choke on a given input file
    // These tests DO NOT test that ClosedXML successfully recognises all the Excel parts or that it can successfully save those parts again.
    [TestFixture]
    public class LoadingTests
    {
        private static IEnumerable<string> TryToLoad =>
            TestHelper.ListResourceFiles(s =>
                    s.Contains(".TryToLoad.") &&
                    !s.Contains(".LO."));

        [TestCaseSource(nameof(TryToLoad))]
        public void CanSuccessfullyLoadFiles(string file)
        {
            TestHelper.LoadFile(file);
        }

        [TestCaseSource(nameof(LOFiles))]
        public void CanSuccessfullyLoadLOFiles(string file)
        {
            TestHelper.LoadFile(file);
        }

        private static IEnumerable<string> LOFiles
        {
            get
            {
                // TODO: unpark all files
                var parkedForLater = new[]
                {
                    "TryToLoad.LO.xlsx.column-style-autofilter.xlsx",
                    "TryToLoad.LO.xlsx.formats.xlsx",
                    "TryToLoad.LO.xlsx.pivot_table.shared-group-field.xlsx",
                    "TryToLoad.LO.xlsx.pivot_table.shared-nested-dategroup.xlsx",
                    "TryToLoad.LO.xlsx.pivottable_bool_field_filter.xlsx",
                    "TryToLoad.LO.xlsx.pivottable_date_field_filter.xlsx",
                    "TryToLoad.LO.xlsx.pivottable_double_field_filter.xlsx",
                    "TryToLoad.LO.xlsx.pivottable_duplicated_member_filter.xlsx",
                    "TryToLoad.LO.xlsx.pivottable_rowcolpage_field_filter.xlsx",
                    "TryToLoad.LO.xlsx.pivottable_string_field_filter.xlsx",
                    "TryToLoad.LO.xlsx.pivottable_tabular_mode.xlsx",
                    "TryToLoad.LO.xlsx.pivot_table_first_header_row.xlsx",
                    "TryToLoad.LO.xlsx.tdf100709.xlsx",
                    "TryToLoad.LO.xlsx.tdf89139_pivot_table.xlsx",
                    "TryToLoad.LO.xlsx.universal-content-strict.xlsx",
                    "TryToLoad.LO.xlsx.universal-content.xlsx",
                    "TryToLoad.LO.xlsx.xf_default_values.xlsx",
                    "TryToLoad.LO.xlsm.pass.CVE-2016-0122-1.xlsm",
                    "TryToLoad.LO.xlsm.tdf111974.xlsm",
                    "TryToLoad.LO.xlsm.vba-user-function.xlsm",
                };

                return TestHelper.ListResourceFiles(s => s.Contains(".LO.") && !parkedForLater.Any(i => s.Contains(i)));
            }
        }

        [Test]
        public void CorrectlyLoadValidationWithSheetReference()
        {
            // Arrange
            var path = TestHelper.GetResourcePath(@"TryToLoad\ValidationWithSheetReference.xlsx");
            using var stream = TestHelper.GetStreamFromResource(path);

            // Act
            using var wb = new XLWorkbook(stream);

            // Assert
            var ws = wb.Worksheet("UI Sheet");
            var B2 = ws.Cell("B2");
            Assert.Multiple(() =>
            {
                Assert.That(B2.GetDataValidation().AllowedValues, Is.EqualTo(XLAllowedValues.List));
                Assert.That(B2.GetDataValidation().Value, Is.EqualTo("$E$1:$E$4"));
            });
            var A2 = ws.Cell("A2");
            Assert.Multiple(() =>
            {
                Assert.That(A2.GetDataValidation().AllowedValues, Is.EqualTo(XLAllowedValues.List));
                Assert.That(A2.GetDataValidation().Value, Is.EqualTo("ValuesSheet!$A$1:$A$4"));
            });
        }

        [Test]
        public void CanLoadAndManipulateFileWithEmptyTable()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\EmptyTable.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();
            var table = ws.Tables.First();
            table.DataRange.InsertRowsBelow(5);
        }

        [Test]
        public void CanLoadDate1904SystemCorrectly()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\Date1904System.xlsx"));
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                var c = ws.Cell("A2");
                Assert.Multiple(() =>
                {
                    Assert.That(c.DataType, Is.EqualTo(XLDataType.DateTime));
                    Assert.That(c.GetDateTime(), Is.EqualTo(new DateTime(2017, 10, 27, 21, 0, 0)));
                });
                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                var c = ws.Cell("A2");
                Assert.That(c.DataType, Is.EqualTo(XLDataType.DateTime));
                Assert.That(c.GetDateTime(), Is.EqualTo(new DateTime(2017, 10, 27, 21, 0, 0)));
                wb.SaveAs(ms);
            }
        }

        [Test]
        public void CanLoadAndSaveFileWithMismatchingSheetIdAndRelId()
        {
            // This file's workbook.xml contains:
            // <x:sheet name="Data" sheetId="13" r:id="rId1" />
            // and the mismatch between the sheetId and r:id can create problems.
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\FileWithMismatchSheetIdAndRelId.xlsx"));
            using var wb = new XLWorkbook(stream);
            using var ms = new MemoryStream();
            wb.SaveAs(ms, true);
        }

        [Test]
        public void CanLoadBasicPivotTable()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LoadPivotTables.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet("PivotTable1");
            var pt = ws.PivotTable("PivotTable1");
            Assert.That(pt.Name, Is.EqualTo("PivotTable1"));

            Assert.That(pt.RowLabels.Count(), Is.EqualTo(1));
            Assert.That(pt.RowLabels.Single().SourceName, Is.EqualTo("Name"));

            Assert.That(pt.ColumnLabels.Count(), Is.EqualTo(1));
            Assert.That(pt.ColumnLabels.Single().SourceName, Is.EqualTo("Month"));

            var pv = pt.Values.Single();
            Assert.That(pv.CustomName, Is.EqualTo("Sum of NumberOfOrders"));
            Assert.That(pv.SourceName, Is.EqualTo("NumberOfOrders"));
        }

        [Test]
        public void CanLoadOrderedPivotTable()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LoadPivotTables.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet("OrderedPivotTable");
            var pt = ws.PivotTable("OrderedPivotTable");

            Assert.That(pt.RowLabels.Single().SortType, Is.EqualTo(XLPivotSortType.Ascending));
            Assert.That(pt.ColumnLabels.Single().SortType, Is.EqualTo(XLPivotSortType.Descending));
        }

        [Test]
        public void CanLoadPivotTableSubtotals()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LoadPivotTables.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet("PivotTableSubtotals");
            var pt = ws.PivotTable("PivotTableSubtotals");

            var subtotals = pt.RowLabels.Get("Group").Subtotals.ToArray();

            Assert.That(subtotals, Is.EquivalentTo(new[]
            {
                XLSubtotalFunction.Average,
                XLSubtotalFunction.Count,
                XLSubtotalFunction.Sum,
            }));
        }

        [Test]
        [Ignore("PT styles will be fixed in a different PR")]
        public void CanLoadPivotTableWithBorder()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\PivotTableWithBorder.xlsx"));
            using var wb = new XLWorkbook(stream);
            var pt = wb.Worksheet(1).PivotTables.PivotTable("PivotTable1");
            var border = pt.RowLabels.Single().StyleFormats.DataValuesFormat.Style.Border;

            Assert.That(border.LeftBorder, Is.EqualTo(XLBorderStyleValues.Thin));
            Assert.That(border.TopBorder, Is.EqualTo(XLBorderStyleValues.Thin));
            Assert.That(border.RightBorder, Is.EqualTo(XLBorderStyleValues.Thin));
            Assert.That(border.BottomBorder, Is.EqualTo(XLBorderStyleValues.Thin));
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
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\FileWithDefaultStyleNameNotInEnglish.xlsx"));
            using var wb = new XLWorkbook(stream);
            using var ms = new MemoryStream();
            wb.SaveAs(ms, true);
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
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LibreOfficeFileWithDates.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();
            foreach (var cell in ws.CellsUsed())
            {
                Assert.That(cell.DataType, Is.EqualTo(XLDataType.DateTime));
            }
        }

        [Test]
        public void CanLoadFileWithImagesWithCorrectAnchorTypes()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();
            Assert.That(ws.Pictures, Has.Count.EqualTo(2));
            Assert.That(ws.Pictures.First().Placement, Is.EqualTo(XLPicturePlacement.FreeFloating));
            Assert.That(ws.Pictures.Skip(1).First().Placement, Is.EqualTo(XLPicturePlacement.Move));

            var ws2 = wb.Worksheets.Skip(1).First();
            Assert.That(ws2.Pictures, Has.Count.EqualTo(1));
            Assert.That(ws2.Pictures.First().Placement, Is.EqualTo(XLPicturePlacement.MoveAndSize));
        }

        [Test]
        public void CanLoadFileWithImagesWithCorrectImageType()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageFormats.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();
            Assert.That(ws.Pictures, Has.Count.EqualTo(1));
            Assert.That(ws.Pictures.First().Format, Is.EqualTo(XLPictureFormat.Jpeg));

            var ws2 = wb.Worksheets.Skip(1).First();
            Assert.That(ws2.Pictures, Has.Count.EqualTo(1));
            Assert.That(ws2.Pictures.First().Format, Is.EqualTo(XLPictureFormat.Png));
        }

        [Test]
        public void CanLoadAndDeduceAnchorsFromExcelGeneratedFile()
        {
            // This file was produced by Excel. It contains 3 images, but the latter 2 were copied from the first.
            // There is actually only 1 embedded image if you inspect the file's internals.
            // Additionally, Excel saves all image anchors as TwoCellAnchor, but uses the EditAs attribute to distinguish the types
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\ExcelProducedWorkbookWithImages.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();
            Assert.That(ws.Pictures, Has.Count.EqualTo(3));

            Assert.That(ws.Picture("Picture 1").Placement, Is.EqualTo(XLPicturePlacement.MoveAndSize));
            Assert.That(ws.Picture("Picture 2").Placement, Is.EqualTo(XLPicturePlacement.Move));
            Assert.That(ws.Picture("Picture 3").Placement, Is.EqualTo(XLPicturePlacement.FreeFloating));

            using var ms = new MemoryStream();
            wb.SaveAs(ms, true);
        }

        [Test]
        public void CanLoadFromTemplate()
        {
            using var tf1 = new TemporaryFile();
            using var tf2 = new TemporaryFile();
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\AllShapes.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                // Save as temporary file
                wb.SaveAs(tf1.Path);
            }

            var workbook = XLWorkbook.OpenFromTemplate(tf1.Path);
            Assert.That(workbook.Worksheets.Any(), Is.True);
            Assert.Throws<InvalidOperationException>(() => workbook.Save());

            workbook.SaveAs(tf2.Path);
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
                using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\EscapedApostrophe.xlsx"));
                using var wb = new XLWorkbook(stream);
                var ws = wb.Worksheets.First();
                title = ws.Name;
            };

            Assert.DoesNotThrow(openWorkbook);
            Assert.That(title, Is.EqualTo("L'E"));
        }

        [Test]
        public void CanRoundTripSheetProtectionForObjects()
        {
            using var book = new XLWorkbook();
            var sheet = book.AddWorksheet("TestSheet");
            sheet.Protect()
                .AllowElement(XLSheetProtectionElements.EditObjects | XLSheetProtectionElements.EditScenarios);

            Assert.That(sheet.Protection.AllowedElements, Is.EqualTo(XLSheetProtectionElements.SelectEverything | XLSheetProtectionElements.EditObjects | XLSheetProtectionElements.EditScenarios));

            using var xlStream = new MemoryStream();
            book.SaveAs(xlStream);

            using var persistedBook = new XLWorkbook(xlStream);
            var persistedSheet = persistedBook.Worksheets.Worksheet(1);

            Assert.That(persistedSheet.Protection.AllowedElements, Is.EqualTo(sheet.Protection.AllowedElements));
        }

        [Test]
        [TestCase("A1*10", 1230)]
        [TestCase("A1/10", 12.3)]
        [TestCase("A1&\" cells\"", "123 cells")]
        [TestCase("A1&\"000\"", "123000")]
        [TestCase("ISNUMBER(A1)", true)]
        [TestCase("ISBLANK(A1)", false)]
        [TestCase("DATE(2018,1,28)", 43128)]
        public void LoadFormulaCachedValue(string formula, object expectedCachedValue)
        {
            using var ms = new MemoryStream();
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
                Assert.Multiple(() =>
                {
                    Assert.That(ws.Cell("A2").NeedsRecalculation, Is.False);
                    Assert.That(ws.Cell("A2").CachedValue, Is.EqualTo(expectedCachedValue));
                });
            }
        }

        [Test]
        public void LoadingOptions()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Misc\Formulas.xlsx"));
            Assert.DoesNotThrow(() =>
            {
                // The value in the file is blank and kept.
                using var wb = new XLWorkbook(stream, new LoadOptions { RecalculateAllFormulas = false });
                Assert.That(wb.Worksheets.Single().Cell("C2").CachedValue, Is.EqualTo(Blank.Value));
            });

            Assert.DoesNotThrow(() =>
            {
                // The value in the file is blank, but recalculation sets it to correct 3.
                using var wb = new XLWorkbook(stream, new LoadOptions { RecalculateAllFormulas = true });
                Assert.That(wb.Worksheets.Single().Cell("C2").CachedValue, Is.EqualTo(3));
            });

            Assert.That(new XLWorkbook(stream, new LoadOptions { Dpi = new Point(30, 14) }).DpiX, Is.EqualTo(30));
            Assert.That(new XLWorkbook(stream, new LoadOptions { Dpi = new Point(30, 14) }).DpiY, Is.EqualTo(14));
        }

        [Test]
        public void CanLoadWorksheetStyle()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\BaseColumnWidth.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet(1);

            Assert.That(ws.Style.Font.FontSize, Is.EqualTo(8));
            Assert.That(ws.Style.Font.FontName, Is.EqualTo("Arial"));
            Assert.That(ws.Cell("A1").Style.Font.FontSize, Is.EqualTo(8));
            Assert.That(ws.Cell("A1").Style.Font.FontName, Is.EqualTo("Arial"));
        }

        [Test]
        public void CanCorrectLoadWorkbookCellWithStringDataType()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\CellWithStringDataType.xlsx"));
            using var wb = new XLWorkbook(stream);
            var cellToCheck = wb.Worksheet(1).Cell("B2");
            Assert.That(cellToCheck.DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(cellToCheck.Value, Is.EqualTo("String with String Data type"));
        }

        [Test]
        public void CanCorrectLoadWorkbookCellsWithDateTimeDataTypeOrFormatting()
        {
            const string expected = "03/14/2012 13:30:55";
            TestHelper.LoadAndAssert(wb =>
            {
                for (int row = 2; row < 18; row++)
                {
                    var cellToCheck = wb.Worksheet(1).Cell(row, 2);
                    Assert.That(cellToCheck.DataType, Is.EqualTo(XLDataType.DateTime), $"Cell B{row} has incorrect DataType");
                    Assert.That(cellToCheck.Value.ToString(CultureInfo.InvariantCulture), Is.EqualTo(expected), $"Cell B{row} value differs");
                }
            }, @"TryToLoad\CellsWithDateTimeDataTypeOrFormatting.xlsx");
        }

        [Test]
        public void CanCorrectLoadWorkbookCellsWithTimeSpanDataTypeOrFormatting()
        {
            string[] expected = Enumerable.Range(0, 10).Select(_ => "13:30:55.2").Concat(new[] { "0:30:55.2" }).ToArray();
            TestHelper.LoadAndAssert(wb =>
            {
                for (int i = 0, row = 2; i < expected.Length; i++, row++)
                {
                    var cellToCheck = wb.Worksheet(1).Cell(row, 2);
                    Assert.That(cellToCheck.DataType, Is.EqualTo(XLDataType.TimeSpan), $"Cell B{row} has incorrect DataType");
                    Assert.That(cellToCheck.Value.ToString(CultureInfo.InvariantCulture), Is.EqualTo(expected[i]), $"Cell B{row} value differs");
                }
            }, @"TryToLoad\CellsWithTimeSpanDataTypeOrFormatting.xlsx");
        }

        [Test]
        public void CanCorrectLoadWorkbookCellsWithDateTimesWithLocalePrefix()
        {
            TestHelper.LoadAndAssert(wb =>
            {
                var ws = wb.Worksheet(1);

                Assert.That(ws.Cell(1, 1).GetFormattedString(), Is.EqualTo("21 January 2019"));
                Assert.That(ws.Cell(2, 1).GetFormattedString(), Is.EqualTo("21-Jan-19"));
                Assert.That(ws.Cell(3, 1).GetFormattedString(), Is.EqualTo("Monday, 21 January 2019"));
                Assert.That(ws.Cell(4, 1).GetFormattedString(), Is.EqualTo("21 Jan 2019"));
            }, @"TryToLoad\CellsWithDateTimeWithLocalePrefix.xlsx");
        }

        [Test]
        public void CanCorrectLoadWorkbookDefaultColumnWidth()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Styles\DefaultStyles.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var defaultColumnWidth = wb.ColumnWidth;
                var pixelWidth = XLHelper.NoCToPixels(defaultColumnWidth, wb.Style.Font, wb);
                Assert.That(defaultColumnWidth, Is.EqualTo(8.43).Within(XLHelper.Epsilon));
                Assert.That(pixelWidth, Is.EqualTo(64));
            }

            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\DefaultColumnWidth.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var defaultColumnWidth = wb.ColumnWidth;
                var pixelWidth = XLHelper.NoCToPixels(defaultColumnWidth, wb.Style.Font, wb);
                Assert.That(defaultColumnWidth, Is.EqualTo(8.5).Within(XLHelper.Epsilon));
                Assert.That(pixelWidth, Is.EqualTo(56));
            }
        }

        [Test]
        public void CanCorrectLoadWorksheetBaseColumnWidth()
        {
            // default calibi font case
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Styles\DefaultStyles.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet(1);
                Assert.Multiple(() =>
                {
                    Assert.That(ws.ColumnWidth, Is.EqualTo(8.43).Within(XLHelper.Epsilon));
                    Assert.That(ws.Column(1).Width, Is.EqualTo(8.43).Within(XLHelper.Epsilon));
                });
            }

            // the worksheet has base column width.
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\BaseColumnWidth.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet(1);
                Assert.Multiple(() =>
                {
                    Assert.That(ws.ColumnWidth, Is.EqualTo(11.17).Within(XLHelper.Epsilon));
                    Assert.That(ws.Column(1).Width, Is.EqualTo(11.17).Within(XLHelper.Epsilon));
                });
            }
        }

        [Test]
        public void CanCorrectLoadWorksheetDefaultColumnWidth()
        {
            // worksheet has default column width.
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\SheetDefaultColumnWidth.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet(1);
            double pixelWidth = XLHelper.NoCToPixels(ws.Column(1).Width, ws.Style.Font, wb);
            Assert.That(ws.ColumnWidth, Is.EqualTo(19.75).Within(XLHelper.Epsilon));
            Assert.That(pixelWidth, Is.EqualTo(163).Within(XLHelper.Epsilon));
        }

        [Test]
        public void CanLoadFileWithInvalidSelectedRanges()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\SelectedRanges\InvalidSelectedRange.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet(1);

            Assert.That(ws.SelectedRanges, Has.Count.EqualTo(2));
            Assert.That(ws.SelectedRanges.First().RangeAddress.ToString(), Is.EqualTo("B2:B2"));
            Assert.That(ws.SelectedRanges.Last().RangeAddress.ToString(), Is.EqualTo("B2:C2"));
        }

        [Test]
        public void CanLoadCellsWithoutReferencesCorrectly()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LO\xlsx\row-index-1-based.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet(1);

            Assert.That(ws.Name, Is.EqualTo("Page 1"));

            var expected = new Dictionary<string, XLCellValue>
            {
                ["A1"] = "Action Plan.Name",
                ["B1"] = "Action Plan.Description",
                ["A2"] = "Jerry",
                ["B2"] = "This is a longer Text.\nSecond line.\nThird line.",
                ["A3"] = Blank.Value,
                ["B3"] = Blank.Value
            };

            foreach (var pair in expected)
                Assert.That(ws.Cell(pair.Key).Value, Is.EqualTo(pair.Value), pair.Key);
        }

        [Test]
        public void CorrectlyLoadThemeColors()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\ThemeColors\inputfile.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet(1);

            var c = ws.Cell("A1");
            var themeColor = c.Style.Fill.BackgroundColor.ThemeColor;
            Assert.Multiple(() =>
            {
                Assert.That(themeColor, Is.EqualTo(XLThemeColor.Accent2));
                Assert.That(wb.Theme.ResolveThemeColor(themeColor).Color.ToHex(), Is.EqualTo("FFED7D31"));
            });

            c = ws.Cell("A2");
            themeColor = c.Style.Fill.BackgroundColor.ThemeColor;
            Assert.Multiple(() =>
            {
                Assert.That(themeColor, Is.EqualTo(XLThemeColor.Accent4));
                Assert.That(wb.Theme.ResolveThemeColor(themeColor).Color.ToHex(), Is.EqualTo("FFFFC000"));
            });

            c = ws.Cell("A3");
            themeColor = c.Style.Fill.BackgroundColor.ThemeColor;
            Assert.Multiple(() =>
            {
                Assert.That(themeColor, Is.EqualTo(XLThemeColor.Accent6));
                Assert.That(wb.Theme.ResolveThemeColor(themeColor).Color.ToHex(), Is.EqualTo("FF70AD47"));
            });
        }

        [Test]
        public void CorrectlyLoadMergedCellsBorder()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\MergedCellsBorder\inputfile.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet(1);

            var c = ws.Cell("B2");
            Assert.Multiple(() =>
            {
                Assert.That(c.Style.Border.TopBorderColor.ColorType, Is.EqualTo(XLColorType.Theme));
                Assert.That(c.Style.Border.TopBorderColor.ThemeColor, Is.EqualTo(XLThemeColor.Accent1));
                Assert.That(c.Style.Border.TopBorderColor.ThemeTint, Is.EqualTo(0.39994506668294322d).Within(XLHelper.Epsilon));
            });
        }

        [Test]
        public void CorrectlyLoadDefaultRowAndColumnStyles()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\RowAndColumnStyles\inputfile.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet(1);

            Assert.Multiple(() =>
            {
                Assert.That(ws.Row(1).Style.Font.FontSize, Is.EqualTo(8));
                Assert.That(ws.Row(2).Style.Font.FontSize, Is.EqualTo(8));
                Assert.That(ws.Column("A").Style.Font.FontSize, Is.EqualTo(8));
            });
        }

        [Test]
        public void EmptyNumberFormatIdTreatedAsGeneral()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\EmptyNumberFormatId.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet(1);

            Assert.That(ws.Cell("A2").Style.NumberFormat.NumberFormatId, Is.EqualTo(XLPredefinedFormat.General));
        }

        [Test]
        public void CanLoadProperties()
        {
            const string author = "TestAuthor";
            const string title = "TestTitle";
            const string subject = "TestSubject";
            const string category = "TestCategory";
            const string keywords = "TestKeywords";
            const string comments = "TestComments";
            const string status = "TestStatus";
            var created = new DateTime(2019, 10, 19, 20, 42, 30);
            var modified = new DateTime(2020, 11, 20, 09, 51, 20);
            const string lastModifiedBy = "TestLastModifiedBy";
            const string company = "TestCompany";
            const string manager = "TestManager";

            using var stream = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.AddWorksheet("sheet1");

                wb.Properties.Author = author;
                wb.Properties.Title = title;
                wb.Properties.Subject = subject;
                wb.Properties.Category = category;
                wb.Properties.Keywords = keywords;
                wb.Properties.Comments = comments;
                wb.Properties.Status = status;
                wb.Properties.Created = created;
                wb.Properties.Modified = modified;
                wb.Properties.LastModifiedBy = lastModifiedBy;
                wb.Properties.Company = company;
                wb.Properties.Manager = manager;

                wb.SaveAs(stream, true);
            }

            stream.Position = 0;

            using (var wb = new XLWorkbook(stream))
            {
                Assert.That(wb.Properties.Author, Is.EqualTo(author));
                Assert.That(wb.Properties.Title, Is.EqualTo(title));
                Assert.That(wb.Properties.Subject, Is.EqualTo(subject));
                Assert.That(wb.Properties.Category, Is.EqualTo(category));
                Assert.That(wb.Properties.Keywords, Is.EqualTo(keywords));
                Assert.That(wb.Properties.Comments, Is.EqualTo(comments));
                Assert.That(wb.Properties.Status, Is.EqualTo(status));
                Assert.That(wb.Properties.Created, Is.EqualTo(created));
                Assert.That(wb.Properties.Modified, Is.EqualTo(modified));
                Assert.That(wb.Properties.LastModifiedBy, Is.EqualTo(lastModifiedBy));
                Assert.That(wb.Properties.Company, Is.EqualTo(company));
                Assert.That(wb.Properties.Manager, Is.EqualTo(manager));
            }
        }

        [Test]
        public void CanLoadEmptyStyles()
        {
            // Stylesheet part exists, but no style collection elements are present
            TestHelper.LoadAndAssert(wb =>
            {
                using var ms = new MemoryStream();
                wb.SaveAs(ms, true);
            }, @"TryToLoad\EmptyStyles.xlsx");
        }

        [Test]
        public void CanLoadInvalidColors()
        {
            // The styles.xml contains two invalid colors: '0' and 'FED+'. Both
            // should be loaded and no exception thrown. The colors are
            // converted using an Excel algorithm.
            TestHelper.LoadAndAssert(wb =>
            {
                var ws = wb.Worksheets.Single();
                Assert.That(ws.Cell("A1").Style.Font.FontColor, Is.EqualTo(XLColor.FromArgb(0xFF000000)));
                Assert.That(ws.Cell("A2").Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromArgb(0xFF000FED)));
            }, @"TryToLoad\InvalidColors.xlsx");
        }

        [Test]
        public void WontCrashOnSheetsWithoutRelId()
        {
            // Some non-Excel producers create workbooks where workbookPart declares
            // sheet with empty r:id, but with name and sheetId. Content of such sheets
            // isn't loaded even if relationship part declares implicit relationship to
            // the worksheets, because workbook has explicit relationships with worksheet
            // part (ISO29500 12.3.23).
            //
            // If excel finds sheet in workbook without r:id, it adds empty sheet with
            // the specified name and so does ClosedXML.
            TestHelper.LoadAndAssert(wb =>
            {
                Assert.That(wb.Worksheets, Has.Count.EqualTo(3));

                // First sheet has r:id, so it keeps content
                Assert.That(wb.Worksheet("Sheet1").Cell("A1").Value, Is.EqualTo("Sheet1"));

                // Second sheet doesn't have r:id, so it is empty after load.
                Assert.That(wb.Worksheet("Sheet without relId").Cell("A1").Value, Is.EqualTo(Blank.Value));

                // Third sheet doesn't have r:id and it contains pivot table that is not loaded.
                var ptSheet = wb.Worksheet("Pivot Sheet without relId");
                Assert.That(ptSheet.Cell("A1").Value, Is.EqualTo(Blank.Value));
                Assert.That(ptSheet.PivotTables.Any(), Is.False);
            }, @"TryToLoad\SheetsWithoutRelId.xlsx");
        }

        [Test]
        public void CanLoadDialogSheet()
        {
            // Workbook can reference multiple different types of sheet, most common is worksheet,
            // but there is also possibility of referencing dialogSheet (basically VBA dialog).
            // dialogSheet is basically obsolete (from Excel 5.0), but still supported. Do not
            // crash when such sheet is encountered. Test file also contains pivot table, because
            // it originally crashed just before pivot table loading.
            TestHelper.LoadAndAssert(wb =>
            {
                // Dialog sheet
                Assert.That(wb.UnsupportedSheets, Has.Count.EqualTo(1));

                // Data and pivot sheets
                Assert.That(wb.Worksheets.Count, Is.EqualTo(2));
                Assert.That(wb.Worksheet("Pivot").PivotTables.Contains("PivotTable1"), Is.True);
            }, @"TryToLoad\DialogSheet.xlsx");
        }
    }
}
