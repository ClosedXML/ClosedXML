using ClosedXML.Excel;
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
                @"Misc\LoadPivotTables.xlsx"
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
    }
}
