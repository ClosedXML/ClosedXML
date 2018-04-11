using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using ClosedXML_Tests.Utils;
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
    }
}