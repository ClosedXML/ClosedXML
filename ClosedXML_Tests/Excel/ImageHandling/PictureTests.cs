using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using NUnit.Framework;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class PictureTests
    {
        [Test]
        public void CanAddPictureFromBitmap()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                using (var resourceStream = Assembly.GetAssembly(typeof(ClosedXML_Examples.BasicTable)).GetManifestResourceStream("ClosedXML_Examples.Resources.SampleImage.jpg"))
                using (var bitmap = Bitmap.FromStream(resourceStream) as Bitmap)
                {
                    var picture = ws.AddPicture(bitmap, "MyPicture")
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .MoveTo(XLMeasure.Create(50, XLMeasureUnit.Pixels), XLMeasure.Create(50, XLMeasureUnit.Pixels))
                        .WithSize(XLMeasure.Create(200, XLMeasureUnit.Pixels), XLMeasure.Create(200, XLMeasureUnit.Pixels));

                    Assert.AreEqual(XLPictureFormat.Jpeg, picture.Format);
                    Assert.AreEqual(200, picture.Width.Value);
                    Assert.AreEqual(200, picture.Height.Value);
                }
            }
        }

        [Test]
        public void CanAddPictureFromStream()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                using (var resourceStream = Assembly.GetAssembly(typeof(ClosedXML_Examples.BasicTable)).GetManifestResourceStream("ClosedXML_Examples.Resources.SampleImage.jpg"))
                {
                    var picture = ws.AddPicture(resourceStream, "MyPicture")
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .MoveTo(XLMeasure.Create(50, XLMeasureUnit.Pixels), XLMeasure.Create(50, XLMeasureUnit.Pixels))
                        .WithSize(XLMeasure.Create(200, XLMeasureUnit.Pixels), XLMeasure.Create(200, XLMeasureUnit.Pixels));

                    Assert.AreEqual(XLPictureFormat.Jpeg, picture.Format);
                    Assert.AreEqual(200, picture.Width.Value);
                    Assert.AreEqual(200, picture.Height.Value);
                }
            }
        }

        [Test]
        public void CanAddPictureFromFile()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                var path = Path.ChangeExtension(Path.GetTempFileName(), "jpg");

                try
                {
                    using (var resourceStream = Assembly.GetAssembly(typeof(ClosedXML_Examples.BasicTable)).GetManifestResourceStream("ClosedXML_Examples.Resources.SampleImage.jpg"))
                    using (var fileStream = File.Create(path))
                    {
                        resourceStream.Seek(0, SeekOrigin.Begin);
                        resourceStream.CopyTo(fileStream);
                        fileStream.Close();
                    }

                    var picture = ws.AddPicture(path)
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .MoveTo(XLMeasure.Create(50, XLMeasureUnit.Pixels), XLMeasure.Create(50, XLMeasureUnit.Pixels));

                    Assert.AreEqual(XLPictureFormat.Jpeg, picture.Format);
                    Assert.AreEqual(400, picture.Width.Value);
                    Assert.AreEqual(400, picture.Height.Value);
                }
                finally
                {
                    if (File.Exists(path))
                        File.Delete(path);
                }
            }
        }

        [Test]
        public void CanAddPictureConcurrentlyFromFile()
        {
            var path = Path.ChangeExtension(Path.GetTempFileName(), "jpg");

            try
            {
                using (var resourceStream = Assembly.GetAssembly(typeof(ClosedXML_Examples.BasicTable)).GetManifestResourceStream("ClosedXML_Examples.Resources.SampleImage.jpg"))
                using (var fileStream = File.Create(path))
                {
                    resourceStream.Seek(0, SeekOrigin.Begin);
                    resourceStream.CopyTo(fileStream);
                    fileStream.Close();
                }

                Parallel.Invoke(() => verifyAddImageFromFile(path), () => verifyAddImageFromFile(path));
            }
            finally
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
        }

        private void verifyAddImageFromFile(string filePath)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                var picture = ws.AddPicture(filePath)
                           .WithPlacement(XLPicturePlacement.FreeFloating)
                           .MoveTo(XLMeasure.Create(50, XLMeasureUnit.Pixels), XLMeasure.Create(50, XLMeasureUnit.Pixels));

                Assert.AreEqual(XLPictureFormat.Jpeg, picture.Format);
                Assert.AreEqual(XLMeasure.Create(400, XLMeasureUnit.Pixels), picture.Width);
                Assert.AreEqual(XLMeasure.Create(50, XLMeasureUnit.Pixels), picture.Top);
            }
        }

        [Test]
        public void CanScaleImage()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                using (var resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Tests.Resource.Images.ImageHandling.png"))
                {
                    var pic = ws.AddPicture(resourceStream, "MyPicture")
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .MoveTo(XLMeasure.Create(50, XLMeasureUnit.Pixels), XLMeasure.Create(50, XLMeasureUnit.Pixels));

                    Assert.AreEqual(252, pic.OriginalWidth.Value);
                    Assert.AreEqual(152, pic.OriginalHeight.Value);
                    Assert.AreEqual(252, pic.Width.Value);
                    Assert.AreEqual(152, pic.Height.Value);

                    pic.ScaleHeight(0.7);
                    pic.ScaleWidth(1.2);

                    Assert.AreEqual(252, pic.OriginalWidth.Value);
                    Assert.AreEqual(152, pic.OriginalHeight.Value);
                    Assert.AreEqual(302, pic.Width.Value);
                    Assert.AreEqual(106, pic.Height.Value);

                    pic.ScaleHeight(0.7);
                    pic.ScaleWidth(1.2);

                    Assert.AreEqual(252, pic.OriginalWidth.Value);
                    Assert.AreEqual(152, pic.OriginalHeight.Value);
                    Assert.AreEqual(362, pic.Width.Value);
                    Assert.AreEqual(74, pic.Height.Value);

                    pic.ScaleHeight(0.8, true);
                    pic.ScaleWidth(1.1, true);

                    Assert.AreEqual(252, pic.OriginalWidth.Value);
                    Assert.AreEqual(152, pic.OriginalHeight.Value);
                    Assert.AreEqual(277, pic.Width.Value);
                    Assert.AreEqual(122, pic.Height.Value);
                }
            }
        }

        [Test]
        public void TestDefaultPictureNames()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Tests.Resource.Images.ImageHandling.png"))
                {
                    ws.AddPicture(stream, XLPictureFormat.Png);
                    stream.Position = 0;

                    ws.AddPicture(stream, XLPictureFormat.Png);
                    stream.Position = 0;

                    ws.AddPicture(stream, XLPictureFormat.Png).Name = "Picture 4";
                    stream.Position = 0;

                    ws.AddPicture(stream, XLPictureFormat.Png);
                    stream.Position = 0;
                }

                Assert.AreEqual("Picture 1", ws.Pictures.Skip(0).First().Name);
                Assert.AreEqual("Picture 2", ws.Pictures.Skip(1).First().Name);
                Assert.AreEqual("Picture 4", ws.Pictures.Skip(2).First().Name);
                Assert.AreEqual("Picture 5", ws.Pictures.Skip(3).First().Name);
            }
        }

        [Test]
        public void TestDefaultIds()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Tests.Resource.Images.ImageHandling.png"))
                {
                    ws.AddPicture(stream, XLPictureFormat.Png);
                    stream.Position = 0;

                    ws.AddPicture(stream, XLPictureFormat.Png);
                    stream.Position = 0;

                    ws.AddPicture(stream, XLPictureFormat.Png).Name = "Picture 4";
                    stream.Position = 0;

                    ws.AddPicture(stream, XLPictureFormat.Png);
                    stream.Position = 0;
                }

                Assert.AreEqual(1, ws.Pictures.Skip(0).First().Id);
                Assert.AreEqual(2, ws.Pictures.Skip(1).First().Id);
                Assert.AreEqual(3, ws.Pictures.Skip(2).First().Id);
                Assert.AreEqual(4, ws.Pictures.Skip(3).First().Id);
            }
        }

        [Test]
        public void XLMarkerTests()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            XLMarker firstMarker = new XLMarker(ws.Cell(1, 10), XLMeasure.Create(100, XLMeasureUnit.Pixels), XLMeasure.Zero);

            Assert.AreEqual(10, firstMarker.ColumnNumber);
            Assert.AreEqual(1, firstMarker.RowNumber);
            Assert.AreEqual(100, firstMarker.X.Value);
            Assert.AreEqual(XLMeasureUnit.Pixels, firstMarker.X.Unit);
            Assert.AreEqual(0, firstMarker.Y.Value);
        }

        [Test]
        public void XLPictureTests()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");

                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Tests.Resource.Images.ImageHandling.png"))
                {
                    var pic = ws.AddPicture(stream, XLPictureFormat.Png, "Image1")
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .MoveTo(XLMeasure.Create(220, XLMeasureUnit.Pixels), XLMeasure.Create(155, XLMeasureUnit.Pixels));

                    Assert.AreEqual(XLPicturePlacement.FreeFloating, pic.Placement);
                    Assert.AreEqual("Image1", pic.Name);
                    Assert.AreEqual(XLPictureFormat.Png, pic.Format);
                    Assert.AreEqual(252, pic.OriginalWidth.Value);
                    Assert.AreEqual(152, pic.OriginalHeight.Value);
                    Assert.AreEqual(252, pic.Width.Value);
                    Assert.AreEqual(152, pic.Height.Value);
                    Assert.AreEqual(220, pic.Left.Value);
                    Assert.AreEqual(155, pic.Top.Value);
                }
            }
        }

        [Test]
        public void CanLoadFileWithImagesAndCopyImagesToNewSheet()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                Assert.AreEqual(2, ws.Pictures.Count);

                var copy = ws.CopyTo("NewSheet");
                Assert.AreEqual(2, copy.Pictures.Count);
            }
        }

        [Test]
        public void CanDeletePictures()
        {
            using (var ms = new MemoryStream())
            {
                int originalCount;

                using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")))
                using (var wb = new XLWorkbook(stream))
                {
                    var ws = wb.Worksheets.First();
                    originalCount = ws.Pictures.Count;
                    ws.Pictures.Delete(ws.Pictures.First());

                    var pictureName = ws.Pictures.First().Name;
                    ws.Pictures.Delete(pictureName);

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheets.First();
                    Assert.AreEqual(originalCount - 2, ws.Pictures.Count);
                }
            }
        }

        [Test]
        public void PictureRenameTests()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet("Images3");
                var picture = ws.Pictures.First();
                Assert.AreEqual("Picture 1", picture.Name);

                picture.Name = "picture 1";
                picture.Name = "pICture 1";
                picture.Name = "Picture 1";

                picture = ws.Pictures.Last();
                picture.Name = "new name";

                Assert.Throws<ArgumentException>(() => picture.Name = "Picture 1");
                Assert.Throws<ArgumentException>(() => picture.Name = "picTURE 1");
            }
        }

        [Test]
        public void HandleDuplicatePictureIdsAcrossWorksheets()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet("Sheet1");
                var ws2 = wb.AddWorksheet("Sheet2");

                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Tests.Resource.Images.ImageHandling.png"))
                {
                    (ws1 as XLWorksheet).AddPicture(stream, "Picture 1", 2);
                    (ws1 as XLWorksheet).AddPicture(stream, "Picture 2", 3);

                    //Internal method - used for loading files
                    var pic = (ws2 as XLWorksheet).AddPicture(stream, "Picture 1", 2)
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .MoveTo(XLMeasure.Create(220, XLMeasureUnit.Pixels), XLMeasure.Create(155, XLMeasureUnit.Pixels)) as XLPicture;

                    var id = pic.Id;

                    pic.Id = id;
                    Assert.AreEqual(id, pic.Id);

                    pic.Id = 3;
                    Assert.AreEqual(3, pic.Id);

                    pic.Id = id;

                    var pic2 = (ws2 as XLWorksheet).AddPicture(stream, "Picture 2", 3)
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .MoveTo(XLMeasure.Create(440, XLMeasureUnit.Pixels), XLMeasure.Create(300, XLMeasureUnit.Pixels)) as XLPicture;
                }
            }
        }

        [Test]
        public void CopyImageSameWorksheet()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");

            IXLPicture original;
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Tests.Resource.Images.ImageHandling.png"))
            {
                original = (ws1 as XLWorksheet).AddPicture(stream, "Picture 1", 2)
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .MoveTo(XLMeasure.Create(220, XLMeasureUnit.Pixels), XLMeasure.Create(155, XLMeasureUnit.Pixels)) as XLPicture;
            }

            var copy = original.Duplicate()
                .MoveTo(XLMeasure.Create(300, XLMeasureUnit.Pixels), XLMeasure.Create(200, XLMeasureUnit.Pixels)) as XLPicture;

            Assert.AreEqual(2, ws1.Pictures.Count());
            Assert.AreEqual(ws1, copy.Worksheet);
            Assert.AreEqual(original.Format, copy.Format);
            Assert.AreEqual(original.Height, copy.Height);
            Assert.AreEqual(original.Placement, copy.Placement);
            Assert.AreEqual(original.TopLeftCell.ToString(), copy.TopLeftCell.ToString());
            Assert.AreEqual(original.Width, copy.Width);
            Assert.AreEqual(original.ImageStream.ToArray(), copy.ImageStream.ToArray(), "Image streams differ");

            Assert.AreEqual(200, copy.Top.Value);
            Assert.AreEqual(300, copy.Left.Value);
            Assert.AreNotEqual(original.Id, copy.Id);
            Assert.AreNotEqual(original.Name, copy.Name);
        }

        [Test]
        public void CopyImageDifferentWorksheets()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            IXLPicture original;
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Tests.Resource.Images.ImageHandling.png"))
            {
                original = (ws1 as XLWorksheet).AddPicture(stream, "Picture 1", 2)
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .MoveTo(XLMeasure.Create(220, XLMeasureUnit.Pixels), XLMeasure.Create(155, XLMeasureUnit.Pixels)) as XLPicture;
            }
            var ws2 = wb.Worksheets.Add("Sheet2");

            var copy = original.CopyTo(ws2);

            Assert.AreEqual(1, ws1.Pictures.Count());
            Assert.AreEqual(1, ws2.Pictures.Count());

            Assert.AreEqual(ws2, copy.Worksheet);

            Assert.AreEqual(original.Format, copy.Format);
            Assert.AreEqual(original.Height, copy.Height);
            Assert.AreEqual(original.Left, copy.Left);
            Assert.AreEqual(original.Name, copy.Name);
            Assert.AreEqual(original.Placement, copy.Placement);
            Assert.AreEqual(original.Top, copy.Top);
            Assert.AreEqual(original.TopLeftCell.ToString(), copy.TopLeftCell.ToString());
            Assert.AreEqual(original.Width, copy.Width);
            Assert.AreEqual(original.ImageStream.ToArray(), copy.ImageStream.ToArray(), "Image streams differ");

            Assert.AreNotEqual(original.Id, copy.Id);
        }

        [Test]
        public void PictureShiftsWhenInsertingRows()
        {
            using (var wb = new XLWorkbook())
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Tests.Resource.Images.ImageHandling.png"))
            {
                var ws = wb.Worksheets.Add("ImageShift");
                var picture = ws.AddPicture(stream, XLPictureFormat.Png, "PngImage")
                    .MoveTo(ws.Cell(5, 2))
                    .WithPlacement(XLPicturePlacement.Move);

                ws.Row(2).InsertRowsBelow(20);

                Assert.AreEqual(25, picture.TopLeftCell.Address.RowNumber);
            }
        }

        [Test]
        public void PictureNotFound()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                Assert.Throws<ArgumentOutOfRangeException>(() => ws.Picture("dummy"));
                Assert.Throws<ArgumentOutOfRangeException>(() => ws.Pictures.Delete("dummy"));
            }
        }
    }
}
