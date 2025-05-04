using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using ClosedXML.Tests.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;

namespace ClosedXML.Tests.Excel.Saving
{
    [TestFixture]
    public class SavingTests
    {
        [Test]
        public void BooleanValueSavesAsZeroOrOne()
        {
            // When a cell evaluates to a boolean value, the text in the XML has to be true/false (lowercase only) or 0/1
            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook();
                var ws = wb.AddWorksheet();
                ws.FirstCell().FormulaA1 = "=TRUE";
                return wb;
            }, @"Other\Formulas\BooleanFormulaValues.xlsx", evaluateFormulae: true);
        }

        [Test]
        public void CanSaveEmptyFile()
        {
            using var ms = new MemoryStream();
            using var wb = new XLWorkbook();
            wb.AddWorksheet("Sheet1");
            wb.SaveAs(ms);
        }

        [Test]
        public void CanSuccessfullySaveFileMultipleTimes()
        {
            using var memoryStream = new MemoryStream();
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");

            // Comments might cause duplicate VmlDrawing Id's - ensure it's tested:
            sheet.Cell(1, 1).GetComment().AddText("abc");

            wb.SaveAs(memoryStream, validate: true);

            for (int i = 1; i <= 3; i++)
            {
                sheet.Cell(i, 1).Value = "test" + i;
                wb.SaveAs(memoryStream, validate: true);
            }
        }

        [Test]
        public void CanEscape_xHHHH_Correctly()
        {
            using var ms = new MemoryStream();
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
                Assert.That(ws.FirstCell().Value, Is.EqualTo("Reserve_TT_A_BLOCAGE_CAG_x6904_2"));
            }
        }

        [Test]
        public void CanSaveFileMultipleTimesAfterDeletingWorksheet()
        {
            // https://github.com/ClosedXML/ClosedXML/issues/435

            using var ms = new MemoryStream();
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
                Assert.That(ws.Name, Is.EqualTo("sheet1"));
                ws.Delete();
                book2.Save();
                book2.Save();
            }
        }

        [Test]
        public void CanSaveAndValidateFileInAnotherCulture()
        {
            string[] cultures = new string[] { "it", "de-AT" };

            foreach (var culture in cultures)
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo(culture);

                using var wb = new XLWorkbook();
                var memoryStream = new MemoryStream();
                var ws = wb.Worksheets.Add("Sheet1");

                wb.SaveAs(memoryStream, true);
            }
        }

        [Test]
        public void NotSaveCachedValueWhenFlagIsFalse()
        {
            using var ms = new MemoryStream();
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

                Assert.That(ws.Cell("A2").CachedValue, Is.EqualTo(Blank.Value));
            }
        }

        [Test]
        public void SaveCachedValueWhenFlagIsTrue()
        {
            using var ms = new MemoryStream();
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

                Assert.Multiple(() =>
                {
                    Assert.That(ws.Cell("A2").CachedValue, Is.EqualTo(1230));

                    Assert.That(ws.Cell("A3").CachedValue, Is.EqualTo("1 230"));
                });
            }
        }

        [Test]
        public void CanSaveAsCopyReadOnlyFile()
        {
            using var original = new TemporaryFile();
            try
            {
                using var copy = new TemporaryFile();
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

                Assert.Multiple(() =>
                {
                    // Assert
                    Assert.That(File.Exists(copy.Path), Is.True);
                    Assert.That(File.GetAttributes(copy.Path).HasFlag(FileAttributes.ReadOnly), Is.False);
                });
            }
            finally
            {
                // Tear down
                File.SetAttributes(original.Path, FileAttributes.Normal);
            }
        }

        [Test]
        public void CanSaveAsOverwriteExistingFile()
        {
            using var existing = new TemporaryFile();
            // Arrange
            File.WriteAllText(existing.Path, "");

            // Act
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.Worksheets.Add("TestSheet");
                wb.SaveAs(existing.Path);
            }

            // Assert
            Assert.That(File.Exists(existing.Path), Is.True);
            Assert.That(new FileInfo(existing.Path).Length, Is.GreaterThan(0));
        }

        [Test]
        public void CannotSaveAsOverwriteExistingReadOnlyFile()
        {
            using var existing = new TemporaryFile();
            try
            {
                // Arrange
                File.WriteAllText(existing.Path, "");
                File.SetAttributes(existing.Path, FileAttributes.ReadOnly);

                // Act
                TestDelegate saveAs = () =>
                {
                    using var wb = new XLWorkbook();
                    var sheet = wb.Worksheets.Add("TestSheet");
                    wb.SaveAs(existing.Path);
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

        [Test]
        public void PageBreaksDontDuplicateAtSaving()
        {
            // https://github.com/ClosedXML/ClosedXML/issues/666

            using var ms = new MemoryStream();
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

                Assert.Multiple(() =>
                {
                    Assert.That(ws.PageSetup.ColumnBreaks, Has.Count.EqualTo(1));
                    Assert.That(ws.PageSetup.RowBreaks, Has.Count.EqualTo(1));
                });
            }
        }

        [Test]
        public void CanSaveFileWithPictureAndComment()
        {
            using var ms = new MemoryStream();
            using var wb = new XLWorkbook();
            using var imageStream = Assembly.GetAssembly(typeof(ClosedXML.Examples.BasicTable)).GetManifestResourceStream("ClosedXML.Examples.Resources.SampleImage.jpg");
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("D4").Value = "Hello world.";

            ws.AddPicture(imageStream, "MyPicture")
                .WithPlacement(XLPicturePlacement.FreeFloating)
                .MoveTo(50, 50)
                .WithSize(200, 200);

            ws.Cell("D4").GetComment().SetVisible().AddText("This is a comment");

            wb.SaveAs(ms);
        }

        [Test]
        public void PreserveChartsWhenSaving()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Charts\PreserveCharts\inputfile.xlsx"));
            using var ms = new MemoryStream();
            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook(stream);
                wb.SaveAs(ms);
                return wb;
            }, @"Other\Charts\PreserveCharts\outputfile.xlsx");
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
            using var tf = new TemporaryFile(Path.ChangeExtension(Path.GetTempFileName(), extension));
            using (var wb = new XLWorkbook())
            {
                wb.Worksheets.Add("Sheet1");
                wb.SaveAs(tf.Path);
            }

            using (var package = SpreadsheetDocument.Open(tf.Path, false))
            {
                Assert.That(package.DocumentType, Is.EqualTo(expectedType));
            }
        }

        [Test]
        public void CanSaveTemplateAsWorkbook()
        {
            // See #1375
            using var template = new TemporaryFile(Path.ChangeExtension(Path.GetTempFileName(), "xltx"));
            using var workbook = new TemporaryFile();
            using (var wb = new XLWorkbook())
            {
                wb.AddWorksheet();
                wb.SaveAs(template.Path);
            }
            using (var wb = new XLWorkbook(template.Path))
            {
                wb.SaveAs(workbook.Path);
            }
            using (var package = SpreadsheetDocument.Open(workbook.Path, false))
            {
                Assert.That(package.DocumentType, Is.EqualTo(SpreadsheetDocumentType.Workbook));
            }
        }

        [Test]
        public void SaveAsWithNoExtensionFails()
        {
            using var tf = new TemporaryFile("FileWithNoExtension");
            using var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet1");
            TestDelegate action = () => wb.SaveAs(tf.Path);

            Assert.Throws<ArgumentException>(action);
        }

        [Test]
        public void SaveAsWithUnsupportedExtensionFails()
        {
            using var tf = new TemporaryFile("FileWithBadExtension.bad");
            using var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet1");
            TestDelegate action = () => wb.SaveAs(tf.Path);

            Assert.Throws<ArgumentException>(action);
        }

        [Test]
        public void SaveCellValueWithLeadingQuotationMarkCorrectly()
        {
            var formulaValue = "=IF(TRUE, 1, 0)";
            var quotedFormulaValue = '\'' + formulaValue;
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                var cell = ws.FirstCell();
                cell.SetValue(quotedFormulaValue);
                Assert.Multiple(() =>
                {
                    Assert.That(cell.HasFormula, Is.False);
                    Assert.That(cell.Value, Is.EqualTo(formulaValue));
                    Assert.That(cell.DataType, Is.EqualTo(XLDataType.Text));
                });
                Assert.That(cell.Style.IncludeQuotePrefix, Is.True);

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                var cell = ws.FirstCell();
                Assert.That(cell.HasFormula, Is.False);
                Assert.Multiple(() =>
                {
                    Assert.That(cell.HasFormula, Is.False);
                    Assert.That(cell.Value, Is.EqualTo(formulaValue));
                    Assert.That(cell.DataType, Is.EqualTo(XLDataType.Text));
                });
                Assert.That(cell.Style.IncludeQuotePrefix, Is.True);
            }
        }

        [Test]
        public void PreserveHeightOfEmptyRowsOnSaving()
        {
            using var ms = new MemoryStream();
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

                    Assert.Multiple(() =>
                    {
                        Assert.That(ws.Row(1).Height, Is.EqualTo(50));
                        Assert.That(ws.Row(2).Height, Is.EqualTo(0));
                        Assert.That(ws.Row(3).Height, Is.EqualTo(20));
                        Assert.That(ws.Row(4).Height, Is.EqualTo(100));
                    });
                }
            }
        }

        [Test]
        public void PreserveWidthOfEmptyColumnsOnSaving()
        {
            using var ms = new MemoryStream();
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

                    Assert.Multiple(() =>
                    {
                        Assert.That(ws.Column(1).Width, Is.EqualTo(ws.ColumnWidth));
                        Assert.That(ws.Column(2).Width, Is.EqualTo(0));
                        Assert.That(ws.Column(3).Width, Is.EqualTo(20));
                        Assert.That(ws.Column(4).Width, Is.EqualTo(100));
                    });
                }
            }
        }

        [Test]
        public void PreserveAlignmentOnSaving()
        {
            using var input = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\HorizontalAlignment.xlsx"));
            using var output = new MemoryStream();
            using (var wb = new XLWorkbook(input))
            {
                wb.SaveAs(output);
            }

            using (var wb = new XLWorkbook(output))
            {
                Assert.That(wb.Worksheets.First().Cell("B1").Style.Alignment.Horizontal, Is.EqualTo(XLAlignmentHorizontalValues.Center));
            }
        }

        [Test]
        public void PreserveMultipleColorScalesOnSaving()
        {
            using var output = new MemoryStream();
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
                Assert.That(cf, Has.Length.EqualTo(2));
                Assert.Multiple(() =>
                {
                    Assert.That(cf[0].ConditionalFormatType, Is.EqualTo(XLConditionalFormatType.ColorScale));
                    Assert.That(cf[0].Colors[1], Is.EqualTo(XLColor.Red));
                    Assert.That(cf[0].ContentTypes[1], Is.EqualTo(XLCFContentType.Minimum));
                    Assert.That(cf[0].Colors[2], Is.EqualTo(XLColor.Green));
                    Assert.That(cf[0].ContentTypes[2], Is.EqualTo(XLCFContentType.Maximum));
                    Assert.That(cf[1].ConditionalFormatType, Is.EqualTo(XLConditionalFormatType.ColorScale));
                    Assert.That(cf[1].Colors[1], Is.EqualTo(XLColor.Alizarin));
                    Assert.That(cf[1].ContentTypes[1], Is.EqualTo(XLCFContentType.Minimum));
                    Assert.That(cf[1].Colors[2], Is.EqualTo(XLColor.Blue));
                    Assert.That(cf[1].ContentTypes[2], Is.EqualTo(XLCFContentType.Maximum));
                });
            }
        }

        [Test]
        public void RemoveExistingInlineStringsIfRequired()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\InlineStrings\inputfile.xlsx"));
            using var ms = new MemoryStream();
            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook(stream);
                var ws = wb.Worksheet(1);

                var numericCells = ws.CellsUsed(c => double.TryParse(c.GetString(), out double _));
                var textCells = ws.CellsUsed(c => !double.TryParse(c.GetString(), out double _));

                foreach (var cell in numericCells)
                {
                    cell.Clear(XLClearOptions.AllFormats);
                    Assert.That(cell.Value.TryConvert(out double val, CultureInfo.CurrentCulture), Is.True);
                    cell.Value = val;
                }

                foreach (var cell in textCells)
                {
                    cell.ShareString = true;
                }

                wb.SaveAs(ms);

                return wb;
            }, @"Other\InlineStrings\outputfile.xlsx");
        }

        [Test]
        public void CanSaveFileWithEmptyFill()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\EmptyFill.xlsx"));
            using var wb = new XLWorkbook(stream);
            using var ms = new MemoryStream();
            Assert.DoesNotThrow(() => wb.SaveAs(ms, false));
        }

        [Test]
        public void CanSaveSingleRowAutoFilter()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\SingleRowAutoFilter.xlsx"));
            using var wb = new XLWorkbook(stream);
            using var ms = new MemoryStream();
            Assert.DoesNotThrow(() => wb.SaveAs(ms, false));
        }

        [Test]
        public void PivotTableWithVeryLongField()
        {
            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook();
                var ws = wb.AddWorksheet();

                var longText = string.Join(" ", Enumerable.Range(0, 40).Select(i => "1234567890"));

                var data = new[]
                {
                    new { Col1 = longText, Col2 = 2}
                };

                var table = ws.FirstCell().InsertTable(data);

                var pvtSheet = wb.AddWorksheet("pvt");

                var pvt = table.CreatePivotTable(pvtSheet.FirstCell(), "PivotTable1");
                pvt.RowLabels.Add("Col1");

                return wb;
            }, @"Other\PivotTableReferenceFiles\LongText\outputfile.xlsx");
        }

        [Test]
        public void CanSaveFileWithVml_NoComments()
        {
            //See #1285
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\FileWithButton.xlsm"));
            using var wb = new XLWorkbook(stream);
            using var ms = new MemoryStream();
            Assert.DoesNotThrow(() => wb.SaveAs(ms));
        }

        [Test]
        public void CanEnableWorkbookFilterPrivacyAndSaveInWorkbook()
        {
            using var ms = new MemoryStream();

            using (var wb = new XLWorkbook())
            {
                wb.AddWorksheet();
                wb.SaveAs(ms, new SaveOptions { FilterPrivacy = true });
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = SpreadsheetDocument.Open(ms, false))
            {
                Assert.That((bool)wb.WorkbookPart.Workbook.WorkbookProperties.FilterPrivacy, Is.True);
            }
        }

        [Test]
        public void WorkbookFilterPrivacyIsNotSetByDefault()
        {
            using var ms = new MemoryStream();

            using (var wb = new XLWorkbook())
            {
                wb.AddWorksheet();
                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = SpreadsheetDocument.Open(ms, false))
            {
                Assert.That(wb.WorkbookPart.Workbook.WorkbookProperties.FilterPrivacy, Is.Null);
            }
        }

        [Test]
        public void WorkbookFilterPrivacyIsReadCorrectly()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\FilterPrivacyEnabledWorkbook.xlsx"));
            using var wb = SpreadsheetDocument.Open(stream, false);
            Assert.That((bool)wb.WorkbookPart.Workbook.WorkbookProperties.FilterPrivacy, Is.True);
        }

        [Test]
        public void CanSaveAsWithDataValidationAfterInsertFirstRowsAboveAndInsertFirstColumnsBefore()
        {
            using var wb = new XLWorkbook();
            using var ms = new MemoryStream();
            var ws = wb.AddWorksheet("WithDataValidation");
            ws.Range("B4:B4").CreateDataValidation().WholeNumber.Between(0, 1);

            ws.Row(1).InsertRowsAbove(1);
            var dv = ws.DataValidations.ToArray();
            Assert.That(dv, Has.Length.EqualTo(1));
            Assert.That(dv[0].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("B5:B5"));

            Assert.DoesNotThrow(() => wb.SaveAs(ms));

            ws.Column(1).InsertColumnsBefore(1);
            dv = ws.DataValidations.ToArray();
            Assert.That(dv, Has.Length.EqualTo(1));
            Assert.That(dv[0].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("C5:C5"));

            Assert.DoesNotThrow(() => wb.SaveAs(ms));
        }

        // https://github.com/ClosedXML/ClosedXML/issues/1606
        [Test]
        public void CanSaveGSheetsFileWithNewComment()
        {
            using var ms = new MemoryStream();
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\GoogleSheets\file1.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();
            ws.Cell(1, 1).CreateComment().AddText("Test");
            Assert.DoesNotThrow(() => wb.SaveAs(ms));
        }

        [Test]
        public void CanSaveFileToDefaultDirectory()
        {
            var filename = $"test-{Guid.NewGuid()}.xlsx";
            try
            {
                using var wb = new XLWorkbook();
                wb.AddWorksheet().FirstCell().SetValue("Hello, world!");
                Assert.DoesNotThrow(() => wb.SaveAs(filename));
            }
            finally
            {
                File.Delete(filename);
            }
        }

        [Test]
        public void CanAddNewPartsInWorkbookWithDuplicateRelIds()
        {
            // Both Sheet1 and drawing have same relIds: rId2
            // We can add a new worksheet even when there are parts with same relId
            TestHelper.LoadModifyAndCompare(
                @"Other\Parts\MultiplePartsHaveNonUniqueRelId-input.xlsx",
                wb => wb.AddWorksheet(),
                @"Other\Parts\MultiplePartsHaveNonUniqueRelId-output.xlsx");
        }

        [Test]
        public void WorksheetWithDrawingCanBeModified()
        {
            // Issue 2080: Drawing was loading the workbook DOM from the worksheet part and
            // the OpenXML SDK was ignoring worksheet changes saved through streaming, but used
            // the eager loaded DOM instead.
            // Saved file doesn't contain shape because it's not yet supported (#1252)
            TestHelper.LoadModifyAndCompare(
                @"Other\Parts\WorksheetWithDrawingCanBeModified-input.xlsx",
                wb =>
                {
                    var ws = wb.Worksheets.Single();
                    ws.Cell("A1").Value = "B";
                },
                @"Other\Parts\WorksheetWithDrawingCanBeModified-output.xlsx");
        }

        [Test]
        public void CorrectlySaveValidationWithSheetReference()
        {
            // When validation with sheet reference loading was first implemented, there was a
            // disconnect between where those validations were being loaded from and where they
            // were being saved to. This led to exceptions being thrown when these validations
            // were loaded/saved multiple times, so this test makes sure that the fix for that
            // issue continues to work by forcing multiple load/save cycles.

            var filename1 = $"test-{Guid.NewGuid()}.xlsx";
            var filename2 = $"test-{Guid.NewGuid()}.xlsx";
            try
            {
                var path = TestHelper.GetResourcePath(@"TryToLoad\ValidationWithSheetReference.xlsx");
                using var stream = TestHelper.GetStreamFromResource(path);

                using var originalWorkbook = new XLWorkbook(stream);
                Assert.DoesNotThrow(() => originalWorkbook.SaveAs(filename1));

                using var workbook1 = new XLWorkbook(filename1);
                Assert.DoesNotThrow(() => workbook1.SaveAs(filename2));

                using var workbook2 = new XLWorkbook(filename2);
                var ws = workbook2.Worksheet("UI Sheet");
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
            finally
            {
                File.Delete(filename1);
                File.Delete(filename2);
            }
        }

        [Test]
        public void FormControlsArePreserved()
        {
            // The sheet contains three form controls: two radio buttons and group box.
            // Form controls are rather complex and this test ensures that the saved
            // file still has VML part (that is the source of truth), drawing part
            // (likely a replacement in a decade or two) and three control parts.
            //
            // Also check that custom text of the form controls is preserved (stored in VML).
            TestHelper.LoadSaveAndCompare(
                @"Other\Shapes\sheet-with-form-controls-input.xlsx",
                @"Other\Shapes\sheet-with-form-controls-output.xlsx");
        }
    }
}
