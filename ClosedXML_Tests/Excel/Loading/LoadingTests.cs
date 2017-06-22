using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using NUnit.Framework;
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
                @"Misc\AllShapes.xlsx"
            };

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

        /// <summary>
        /// For non-English locales, the default style ("Normal" in English) can be
        /// another piece of text (e.g. Обычный in Russian).
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
                    Assert.AreEqual(XLCellValues.DateTime, cell.DataType);
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
    }
}
