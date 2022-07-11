using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using ClosedXML.Tests.Utils;
using NUnit.Framework;
using System;
using System.Collections.Generic;
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
        public void CanLoadAndManipulateFileWithEmptyTable()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\EmptyTable.xlsx")))
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
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\Date1904System.xlsx")))
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
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\FileWithMismatchSheetIdAndRelId.xlsx")))
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
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LoadPivotTables.xlsx")))
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
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LoadPivotTables.xlsx")))
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
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LoadPivotTables.xlsx")))
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
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\FileWithDefaultStyleNameNotInEnglish.xlsx")))
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
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LibreOfficeFileWithDates.xlsx")))
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
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\ExcelProducedWorkbookWithImages.xlsx")))
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
                using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\AllShapes.xlsx")))
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
                using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\EscapedApostrophe.xlsx")))
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
                    .AllowElement(XLSheetProtectionElements.EditObjects | XLSheetProtectionElements.EditScenarios);

                Assert.AreEqual(XLSheetProtectionElements.SelectEverything | XLSheetProtectionElements.EditObjects | XLSheetProtectionElements.EditScenarios, sheet.Protection.AllowedElements);

                using (var xlStream = new MemoryStream())
                {
                    book.SaveAs(xlStream);

                    using (var persistedBook = new XLWorkbook(xlStream))
                    {
                        var persistedSheet = persistedBook.Worksheets.Worksheet(1);

                        Assert.AreEqual(sheet.Protection.AllowedElements, persistedSheet.Protection.AllowedElements);
                    }
                }
            }
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
                    Assert.IsFalse(ws.Cell("A2").NeedsRecalculation);
                    Assert.AreEqual(expectedCachedValue, ws.Cell("A2").CachedValue);
                }
            }
        }

        [Test]
        public void LoadingOptions()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\ExternalLinks\WorkbookWithExternalLink.xlsx")))
            {
                Assert.DoesNotThrow(() => new XLWorkbook(stream, new LoadOptions { RecalculateAllFormulas = false }));
                Assert.Throws<NotImplementedException>(() => new XLWorkbook(stream, new LoadOptions { RecalculateAllFormulas = true }));

                Assert.AreEqual(XLEventTracking.Disabled, new XLWorkbook(stream, new LoadOptions { EventTracking = XLEventTracking.Disabled }).EventTracking);
                Assert.AreEqual(XLEventTracking.Enabled, new XLWorkbook(stream, new LoadOptions { EventTracking = XLEventTracking.Enabled }).EventTracking);
            }
        }

        [Test]
        public void CanLoadWorksheetStyle()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\BaseColumnWidth.xlsx")))
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
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\CellWithStringDataType.xlsx")))
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

        [Test]
        public void CanLoadCellsWithoutReferencesCorrectly()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LO\xlsx\row-index-1-based.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet(1);

                Assert.AreEqual("Page 1", ws.Name);

                var expected = new Dictionary<string, string>()
                {
                    ["A1"] = "Action Plan.Name",
                    ["B1"] = "Action Plan.Description",
                    ["A2"] = "Jerry",
                    ["B2"] = "This is a longer Text.\nSecond line.\nThird line.",
                    ["A3"] = "",
                    ["B3"] = ""
                };

                foreach (var pair in expected)
                    Assert.AreEqual(pair.Value, ws.Cell(pair.Key).GetString(), pair.Key);
            }
        }

        [Test]
        public void CorrectlyLoadMergedCellsBorder()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\MergedCellsBorder\inputfile.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet(1);

                var c = ws.Cell("B2");
                Assert.AreEqual(XLColorType.Theme, c.Style.Border.TopBorderColor.ColorType);
                Assert.AreEqual(XLThemeColor.Accent1, c.Style.Border.TopBorderColor.ThemeColor);
                Assert.AreEqual(0.39994506668294322d, c.Style.Border.TopBorderColor.ThemeTint, XLHelper.Epsilon);
            }
        }

        [Test]
        public void CorrectlyLoadDefaultRowAndColumnStyles()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\RowAndColumnStyles\inputfile.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet(1);

                Assert.AreEqual(8, ws.Row(1).Style.Font.FontSize);
                Assert.AreEqual(8, ws.Row(2).Style.Font.FontSize);
                Assert.AreEqual(8, ws.Column("A").Style.Font.FontSize);
            }
        }

        [Test]
        public void EmptyNumberFormatIdTreatedAsGeneral()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\EmptyNumberFormatId.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet(1);

                Assert.AreEqual(XLPredefinedFormat.General, ws.Cell("A2").Style.NumberFormat.NumberFormatId);
            }
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

            using (var stream = new MemoryStream())
            {
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
                    Assert.AreEqual(author, wb.Properties.Author);
                    Assert.AreEqual(title, wb.Properties.Title);
                    Assert.AreEqual(subject, wb.Properties.Subject);
                    Assert.AreEqual(category, wb.Properties.Category);
                    Assert.AreEqual(keywords, wb.Properties.Keywords);
                    Assert.AreEqual(comments, wb.Properties.Comments);
                    Assert.AreEqual(status, wb.Properties.Status);
                    Assert.AreEqual(created, wb.Properties.Created);
                    Assert.AreEqual(modified, wb.Properties.Modified);
                    Assert.AreEqual(lastModifiedBy, wb.Properties.LastModifiedBy);
                    Assert.AreEqual(company, wb.Properties.Company);
                    Assert.AreEqual(manager, wb.Properties.Manager);
                }
            }
        }
    }
}
