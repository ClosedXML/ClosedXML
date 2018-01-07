using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ClosedXML_Tests.Excel.Saving
{
    [TestFixture]
    public class SavingTests
    {
        private string _tempFolder;
        private List<string> _tempFiles;

        [SetUp]
        public void Setup()
        {
            _tempFolder = Path.GetTempPath();
            _tempFiles = new List<string>();
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
            // Arrange
            string id = Guid.NewGuid().ToString();
            string original = string.Format("{0}original{1}.xlsx", _tempFolder, id);
            string copy = string.Format("{0}copy_of_{1}.xlsx", _tempFolder, id);

            using (var wb = new XLWorkbook())
            {
                var sheet = wb.Worksheets.Add("TestSheet");
                wb.SaveAs(original);
                _tempFiles.Add(original);
            }
            System.IO.File.SetAttributes(original, FileAttributes.ReadOnly);

            // Act
            using (var wb = new XLWorkbook(original))
            {
                wb.SaveAs(copy);
                _tempFiles.Add(copy);
            }

            // Assert
            Assert.IsTrue(System.IO.File.Exists(copy));
            Assert.IsFalse(System.IO.File.GetAttributes(copy).HasFlag(FileAttributes.ReadOnly));
        }

        [Test]
        public void CanSaveAsOverwriteExistingFile()
        {
            // Arrange
            string id = Guid.NewGuid().ToString();
            string existing = string.Format("{0}existing{1}.xlsx", _tempFolder, id);

            System.IO.File.WriteAllText(existing, "");
            _tempFiles.Add(existing);

            // Act
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.Worksheets.Add("TestSheet");
                wb.SaveAs(existing);
            }

            // Assert
            Assert.IsTrue(System.IO.File.Exists(existing));
            Assert.Greater(new System.IO.FileInfo(existing).Length, 0);
        }


        [Test]
        public void CannotSaveAsOverwriteExistingReadOnlyFile()
        {
            // Arrange
            string id = Guid.NewGuid().ToString();
            string existing = string.Format("{0}existing{1}.xlsx", _tempFolder, id);

            System.IO.File.WriteAllText(existing, "");
            _tempFiles.Add(existing);
            System.IO.File.SetAttributes(existing, FileAttributes.ReadOnly);

            // Act
            TestDelegate saveAs = () =>
            {
                using (var wb = new XLWorkbook())
                {
                    var sheet = wb.Worksheets.Add("TestSheet");
                    wb.SaveAs(existing);
                }
            };

            // Assert
            Assert.Throws(typeof(UnauthorizedAccessException), saveAs);
        }


        [TearDown]
        public void DeleteTempFiles()
        {
            foreach (var fileName in _tempFiles)
            {
                try
                {
                    System.IO.File.Delete(fileName);
                }
                catch
                { }
            }
            _tempFiles.Clear();
        }
    }
}
