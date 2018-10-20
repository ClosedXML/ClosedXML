using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using ClosedXML_Tests.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;

namespace ClosedXML_Tests.Excel.Saving
{
    [TestFixture]
    public class SavingTests
    {
        [Test]
        public void CanSaveEmptyFile()
        {
            using (var ms = new MemoryStream())
            using (var wb = new XLWorkbook())
            {
                wb.AddWorksheet("Sheet1");
                wb.SaveAs(ms);
            }
        }

        [Test]
        public void CanSuccessfullySaveFileMultipleTimes()
        {
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.Worksheets.Add("TestSheet");

                // Comments might cause duplicate VmlDrawing Id's - ensure it's tested:
                sheet.Cell(1, 1).Comment.AddText("abc");

                var memoryStream = new MemoryStream();
                wb.SaveAs(memoryStream, true);

                for (int i = 1; i <= 3; i++)
                {
                    sheet.Cell(i, 1).Value = "test" + i;
                    wb.SaveAs(memoryStream, true);
                }

                memoryStream.Close();
                memoryStream.Dispose();
            }
        }

        [Test]
        public void CanEscape_xHHHH_Correctly()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet("Sheet1");
                    ws.FirstCell().Value = "Reserve_TT_A_BLOCAGE_CAG_x6904_2";
                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheets.First();
                    Assert.AreEqual("Reserve_TT_A_BLOCAGE_CAG_x6904_2", ws.FirstCell().Value);
                }
            }
        }

        [Test]
        public void CanSaveFileMultipleTimesAfterDeletingWorksheet()
        {
            // https://github.com/ClosedXML/ClosedXML/issues/435

            using (var ms = new MemoryStream())
            {
                using (XLWorkbook book1 = new XLWorkbook())
                {
                    book1.AddWorksheet("sheet1");
                    book1.AddWorksheet("sheet2");

                    book1.SaveAs(ms);
                }
                ms.Position = 0;

                using (XLWorkbook book2 = new XLWorkbook(ms))
                {
                    var ws = book2.Worksheet(1);
                    Assert.AreEqual("sheet1", ws.Name);
                    ws.Delete();
                    book2.Save();
                    book2.Save();
                }
            }
        }

        [Test]
        public void CanSaveAndValidateFileInAnotherCulture()
        {
            string[] cultures = new string[] { "it", "de-AT" };

            foreach (var culture in cultures)
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo(culture);

                using (var wb = new XLWorkbook())
                {
                    var memoryStream = new MemoryStream();
                    var ws = wb.Worksheets.Add("Sheet1");

                    wb.SaveAs(memoryStream, true);
                }
            }
        }

        [Test]
        public void NotSaveCachedValueWhenFlagIsFalse()
        {
            using (var ms = new MemoryStream())
            {
                using (XLWorkbook book1 = new XLWorkbook())
                {
                    var sheet = book1.AddWorksheet("sheet1");
                    sheet.Cell("A1").Value = 123;
                    sheet.Cell("A2").FormulaA1 = "A1*10";
                    book1.RecalculateAllFormulas();
                    var options = new SaveOptions { EvaluateFormulasBeforeSaving = false };

                    book1.SaveAs(ms, options);
                }
                ms.Position = 0;

                using (XLWorkbook book2 = new XLWorkbook(ms))
                {
                    var ws = book2.Worksheet(1);

                    Assert.IsNull(ws.Cell("A2").ValueCached);
                    Assert.IsNull(ws.Cell("A2").CachedValue);
                }
            }
        }

        [Test]
        public void SaveCachedValueWhenFlagIsTrue()
        {
            using (var ms = new MemoryStream())
            {
                using (XLWorkbook book1 = new XLWorkbook())
                {
                    var sheet = book1.AddWorksheet("sheet1");
                    sheet.Cell("A1").Value = 123;
                    sheet.Cell("A2").FormulaA1 = "A1*10";
                    sheet.Cell("A3").FormulaA1 = "TEXT(A2, \"# ###\")";
                    var options = new SaveOptions { EvaluateFormulasBeforeSaving = true };

                    book1.SaveAs(ms, options);
                }
                ms.Position = 0;

                using (XLWorkbook book2 = new XLWorkbook(ms))
                {
                    var ws = book2.Worksheet(1);

                    Assert.AreEqual("1230", ws.Cell("A2").ValueCached);
                    Assert.AreEqual(1230, ws.Cell("A2").CachedValue);

                    Assert.AreEqual("1 230", ws.Cell("A3").ValueCached);
                    Assert.AreEqual("1 230", ws.Cell("A3").CachedValue);
                }
            }
        }

        [Test]
        public void CanSaveAsCopyReadOnlyFile()
        {
            using (var original = new TemporaryFile())
            {
                try
                {
                    using (var copy = new TemporaryFile())
                    {
                        // Arrange
                        using (var wb = new XLWorkbook())
                        {
                            var sheet = wb.Worksheets.Add("TestSheet");
                            wb.SaveAs(original.Path);
                        }
                        File.SetAttributes(original.Path, FileAttributes.ReadOnly);

                        // Act
                        using (var wb = new XLWorkbook(original.Path))
                        {
                            wb.SaveAs(copy.Path);
                        }

                        // Assert
                        Assert.IsTrue(File.Exists(copy.Path));
                        Assert.IsFalse(File.GetAttributes(copy.Path).HasFlag(FileAttributes.ReadOnly));
                    }
                }
                finally
                {
                    // Tear down
                    File.SetAttributes(original.Path, FileAttributes.Normal);
                }
            }
        }

        [Test]
        public void CanSaveAsOverwriteExistingFile()
        {
            using (var existing = new TemporaryFile())
            {
                // Arrange
                File.WriteAllText(existing.Path, "");

                // Act
                using (var wb = new XLWorkbook())
                {
                    var sheet = wb.Worksheets.Add("TestSheet");
                    wb.SaveAs(existing.Path);
                }

                // Assert
                Assert.IsTrue(File.Exists(existing.Path));
                Assert.Greater(new FileInfo(existing.Path).Length, 0);
            }
        }

        [Test]
        public void CannotSaveAsOverwriteExistingReadOnlyFile()
        {
            using (var existing = new TemporaryFile())
            {
                try
                {
                    // Arrange
                    File.WriteAllText(existing.Path, "");
                    File.SetAttributes(existing.Path, FileAttributes.ReadOnly);

                    // Act
                    TestDelegate saveAs = () =>
                    {
                        using (var wb = new XLWorkbook())
                        {
                            var sheet = wb.Worksheets.Add("TestSheet");
                            wb.SaveAs(existing.Path);
                        }
                    };

                    // Assert
                    Assert.Throws(typeof(UnauthorizedAccessException), saveAs);
                }
                finally
                {
                    // Tear down
                    File.SetAttributes(existing.Path, FileAttributes.Normal);
                }
            }
        }

        [Test]
        public void PageBreaksDontDuplicateAtSaving()
        {
            // https://github.com/ClosedXML/ClosedXML/issues/666

            using (var ms = new MemoryStream())
            {
                using (var wb1 = new XLWorkbook())
                {
                    var ws = wb1.Worksheets.Add("Page Breaks");
                    ws.PageSetup.PrintAreas.Add("A1:D5");
                    ws.PageSetup.AddHorizontalPageBreak(2);
                    ws.PageSetup.AddVerticalPageBreak(2);
                    wb1.SaveAs(ms);
                    wb1.Save();
                }
                using (var wb2 = new XLWorkbook(ms))
                {
                    var ws = wb2.Worksheets.First();

                    Assert.AreEqual(1, ws.PageSetup.ColumnBreaks.Count);
                    Assert.AreEqual(1, ws.PageSetup.RowBreaks.Count);
                }
            }
        }

        [Test]
        public void CanSaveFileWithPictureAndComment()
        {
            using (var ms = new MemoryStream())
            using (var wb = new XLWorkbook())
            using (var resourceStream = Assembly.GetAssembly(typeof(ClosedXML_Examples.BasicTable)).GetManifestResourceStream("ClosedXML_Examples.Resources.SampleImage.jpg"))
            using (var bitmap = Bitmap.FromStream(resourceStream) as Bitmap)
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell("D4").Value = "Hello world.";

                ws.AddPicture(bitmap, "MyPicture")
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .MoveTo(50, 50)
                    .WithSize(200, 200);

                ws.Cell("D4").Comment.SetVisible().AddText("This is a comment");

                wb.SaveAs(ms);
            }
        }

        [Test]
        public void PreserveChartsWhenSaving()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Charts\PreserveCharts\inputfile.xlsx")))
            using (var ms = new MemoryStream())
            {
                TestHelper.CreateAndCompare(() =>
                {
                    var wb = new XLWorkbook(stream);
                    wb.SaveAs(ms);
                    return wb;
                }, @"Other\Charts\PreserveCharts\outputfile.xlsx");
            }
        }

        [Test]
        public void DeletingAllPicturesRemovesDrawingPart()
        {
            TestHelper.CreateAndCompare(() =>
            {
                var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx"));
                var wb = new XLWorkbook(stream);
                foreach (var ws in wb.Worksheets)
                {
                    var pictureNames = ws.Pictures.Select(pic => pic.Name).ToArray();
                    foreach (var name in pictureNames)
                        ws.Pictures.Delete(name);
                }

                return wb;
            }, @"Other\Drawings\NoDrawings\outputfile.xlsx");
        }

        [Test]
        [TestCase("xlsx", SpreadsheetDocumentType.Workbook)]
        [TestCase("xlsm", SpreadsheetDocumentType.MacroEnabledWorkbook)]
        [TestCase("xltx", SpreadsheetDocumentType.Template)]
        [TestCase("xltm", SpreadsheetDocumentType.MacroEnabledTemplate)]
        public void SavesAsProperSpreadsheetDocumentType(string extension, SpreadsheetDocumentType expectedType)
        {
            using (var tf = new TemporaryFile(Path.ChangeExtension(Path.GetTempFileName(), extension)))
            {
                using (var wb = new XLWorkbook())
                {
                    wb.Worksheets.Add("Sheet1");
                    wb.SaveAs(tf.Path);
                }

                using (var package = SpreadsheetDocument.Open(tf.Path, false))
                {
                    Assert.AreEqual(expectedType, package.DocumentType);
                }
            }
        }

        [Test]
        public void SaveAsWithNoExtensionFails()
        {
            using (var tf = new TemporaryFile("FileWithNoExtension"))
            using (var wb = new XLWorkbook())
            {
                wb.Worksheets.Add("Sheet1");
                TestDelegate action = () => wb.SaveAs(tf.Path);

                Assert.Throws<ArgumentException>(action);
            }
        }

        [Test]
        public void SaveAsWithUnsupportedExtensionFails()
        {
            using (var tf = new TemporaryFile("FileWithBadExtension.bad"))
            using (var wb = new XLWorkbook())
            {
                wb.Worksheets.Add("Sheet1");
                TestDelegate action = () => wb.SaveAs(tf.Path);

                Assert.Throws<ArgumentException>(action);
            }
        }

        [Test]
        public void SaveCellValueWithLeadingQuotationMarkCorrectly()
        {
            var quotedFormulaValue = "'=IF(TRUE, 1, 0)";
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet("Sheet1");
                    var cell = ws.FirstCell();
                    cell.SetValue(quotedFormulaValue);
                    Assert.IsFalse(cell.HasFormula);
                    Assert.AreEqual(quotedFormulaValue, cell.Value);

                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheets.First();
                    var cell = ws.FirstCell();
                    Assert.IsFalse(cell.HasFormula);
                    Assert.AreEqual(quotedFormulaValue, cell.Value);
                }
            }
        }

        [Test]
        public void PreserveHeightOfEmptyRowsOnSaving()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet("Sheet1");
                    ws.RowHeight = 50;
                    ws.Row(2).Height = 0;
                    ws.Row(3).Height = 20;
                    ws.Row(4).Height = 100;

                    ws.CopyTo("Sheet2");
                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    foreach (var sheetName in new[] { "Sheet1", "Sheet2" })
                    {
                        var ws = wb.Worksheet(sheetName);

                        Assert.AreEqual(50, ws.Row(1).Height);
                        Assert.AreEqual(0, ws.Row(2).Height);
                        Assert.AreEqual(20, ws.Row(3).Height);
                        Assert.AreEqual(100, ws.Row(4).Height);
                    }
                }
            }
        }

        [Test]
        public void PreserveWidthOfEmptyColumnsOnSaving()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet("Sheet1");
                    ws.Column(2).Width = 0;
                    ws.Column(3).Width = 20;
                    ws.Column(4).Width = 100;

                    ws.CopyTo("Sheet2");
                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    foreach (var sheetName in new[] { "Sheet1", "Sheet2" })
                    {
                        var ws = wb.Worksheet(sheetName);

                        Assert.AreEqual(ws.ColumnWidth, ws.Column(1).Width);
                        Assert.AreEqual(0, ws.Column(2).Width);
                        Assert.AreEqual(20, ws.Column(3).Width);
                        Assert.AreEqual(100, ws.Column(4).Width);
                    }
                }
            }
        }

        [Test]
        public void PreserveAlignmentOnSaving()
        {
            using (var input = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\HorizontalAlignment.xlsx")))
            using (var output = new MemoryStream())
            {
                using (var wb = new XLWorkbook(input))
                {
                    wb.SaveAs(output);
                }

                using (var wb = new XLWorkbook(output))
                {
                    Assert.AreEqual(XLAlignmentHorizontalValues.Center, wb.Worksheets.First().Cell("B1").Style.Alignment.Horizontal);
                }
            }
        }

        [Test]
        public void PreserveMultipleColorScalesOnSaving()
        {
            using (var output = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var sheet = wb.Worksheets.Add("test");
                    sheet.Column(1).AddConditionalFormat().ColorScale().LowestValue(XLColor.Red)
                        .HighestValue(XLColor.Green);

                    sheet.Column(2).AddConditionalFormat().ColorScale().LowestValue(XLColor.Alizarin)
                        .HighestValue(XLColor.Blue);

                    wb.SaveAs(output);
                }

                using (var wb = new XLWorkbook(output))
                {
                    var sheet = wb.Worksheets.First();
                    var cf = sheet.ConditionalFormats
                        .OrderBy(x => x.Range.RangeAddress.FirstAddress.ColumnNumber)
                        .ToArray();
                    Assert.AreEqual(2, cf.Length);
                    Assert.AreEqual(XLConditionalFormatType.ColorScale, cf[0].ConditionalFormatType);
                    Assert.AreEqual(XLColor.Red, cf[0].Colors[1]);
                    Assert.AreEqual(XLCFContentType.Minimum, cf[0].ContentTypes[1]);
                    Assert.AreEqual(XLColor.Green, cf[0].Colors[2]);
                    Assert.AreEqual(XLCFContentType.Maximum, cf[0].ContentTypes[2]);
                    Assert.AreEqual(XLConditionalFormatType.ColorScale, cf[1].ConditionalFormatType);
                    Assert.AreEqual(XLColor.Alizarin, cf[1].Colors[1]);
                    Assert.AreEqual(XLCFContentType.Minimum, cf[1].ContentTypes[1]);
                    Assert.AreEqual(XLColor.Blue, cf[1].Colors[2]);
                    Assert.AreEqual(XLCFContentType.Maximum, cf[1].ContentTypes[2]);
                }
            }
        }
    }
}
