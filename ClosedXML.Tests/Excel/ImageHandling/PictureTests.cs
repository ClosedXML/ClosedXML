using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using NUnit.Framework;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace ClosedXML.Tests
{
    [TestFixture]
    public class PictureTests
    {
        [Test]
        public void CanAddPictureFromStream()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            using var resourceStream = Assembly.GetAssembly(typeof(ClosedXML.Examples.BasicTable)).GetManifestResourceStream("ClosedXML.Examples.Resources.SampleImage.jpg");
            var picture = ws.AddPicture(resourceStream, "MyPicture")
                .WithPlacement(XLPicturePlacement.FreeFloating)
                .MoveTo(50, 50)
                .WithSize(200, 200);

            Assert.Multiple(() =>
            {
                Assert.That(picture.Format, Is.EqualTo(XLPictureFormat.Jpeg));
                Assert.That(picture.Width, Is.EqualTo(200));
                Assert.That(picture.Height, Is.EqualTo(200));
            });
        }

        [Test]
        public void CanAddPictureFromFile()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            var path = Path.ChangeExtension(Path.GetTempFileName(), "jpg");

            try
            {
                using (var resourceStream = Assembly.GetAssembly(typeof(ClosedXML.Examples.BasicTable)).GetManifestResourceStream("ClosedXML.Examples.Resources.SampleImage.jpg"))
                using (var fileStream = File.Create(path))
                {
                    resourceStream.Seek(0, SeekOrigin.Begin);
                    resourceStream.CopyTo(fileStream);
                    fileStream.Close();
                }

                var picture = ws.AddPicture(path)
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .MoveTo(50, 50);

                Assert.Multiple(() =>
                {
                    Assert.That(picture.Format, Is.EqualTo(XLPictureFormat.Jpeg));
                    Assert.That(picture.Width, Is.EqualTo(400));
                    Assert.That(picture.Height, Is.EqualTo(400));
                });
            }
            finally
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
        }

        [Test]
        public void CanAddPictureConcurrentlyFromFile()
        {
            var path = Path.ChangeExtension(Path.GetTempFileName(), "jpg");

            try
            {
                using (var resourceStream = Assembly.GetAssembly(typeof(ClosedXML.Examples.BasicTable)).GetManifestResourceStream("ClosedXML.Examples.Resources.SampleImage.jpg"))
                using (var fileStream = File.Create(path))
                {
                    resourceStream.Seek(0, SeekOrigin.Begin);
                    resourceStream.CopyTo(fileStream);
                    fileStream.Close();
                }

                Parallel.Invoke(() => VerifyAddImageFromFile(path), () => VerifyAddImageFromFile(path));
            }
            finally
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
        }

        private static void VerifyAddImageFromFile(string filePath)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            var picture = ws.AddPicture(filePath)
                .WithPlacement(XLPicturePlacement.FreeFloating)
                .MoveTo(50, 50);

            // Do not Wrap inside an Assert.Multiple, will create an error due to parallel invoke.
#pragma warning disable NUnit2045
            Assert.That(picture.Format, Is.EqualTo(XLPictureFormat.Jpeg));
            Assert.That(picture.Width, Is.EqualTo(400));
            Assert.That(picture.Top, Is.EqualTo(50));
#pragma warning restore NUnit2045
        }

        [Test]
        public void CanScaleImage()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            using var resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png");
            var pic = ws.AddPicture(resourceStream, "MyPicture")
                .WithPlacement(XLPicturePlacement.FreeFloating)
                .MoveTo(50, 50);

            Assert.Multiple(() =>
            {
                Assert.That(pic.OriginalWidth, Is.EqualTo(252));
                Assert.That(pic.OriginalHeight, Is.EqualTo(152));
                Assert.That(pic.Width, Is.EqualTo(252));
                Assert.That(pic.Height, Is.EqualTo(152));
            });

            pic.ScaleHeight(0.7);
            pic.ScaleWidth(1.2);

            Assert.Multiple(() =>
            {
                Assert.That(pic.OriginalWidth, Is.EqualTo(252));
                Assert.That(pic.OriginalHeight, Is.EqualTo(152));
                Assert.That(pic.Width, Is.EqualTo(302));
                Assert.That(pic.Height, Is.EqualTo(106));
            });

            pic.ScaleHeight(0.7);
            pic.ScaleWidth(1.2);

            Assert.Multiple(() =>
            {
                Assert.That(pic.OriginalWidth, Is.EqualTo(252));
                Assert.That(pic.OriginalHeight, Is.EqualTo(152));
                Assert.That(pic.Width, Is.EqualTo(362));
                Assert.That(pic.Height, Is.EqualTo(74));
            });

            pic.ScaleHeight(0.8, true);
            pic.ScaleWidth(1.1, true);

            Assert.Multiple(() =>
            {
                Assert.That(pic.OriginalWidth, Is.EqualTo(252));
                Assert.That(pic.OriginalHeight, Is.EqualTo(152));
                Assert.That(pic.Width, Is.EqualTo(277));
                Assert.That(pic.Height, Is.EqualTo(122));
            });
        }

        [Test]
        public void TestDefaultPictureNames()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png"))
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

            Assert.Multiple(() =>
            {
                Assert.That(ws.Pictures.Skip(0).First().Name, Is.EqualTo("Picture 1"));
                Assert.That(ws.Pictures.Skip(1).First().Name, Is.EqualTo("Picture 2"));
                Assert.That(ws.Pictures.Skip(2).First().Name, Is.EqualTo("Picture 4"));
                Assert.That(ws.Pictures.Skip(3).First().Name, Is.EqualTo("Picture 5"));
            });
        }

        [Test]
        public void TestDefaultIds()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png"))
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

            Assert.Multiple(() =>
            {
                Assert.That(ws.Pictures.Skip(0).First().Id, Is.EqualTo(1));
                Assert.That(ws.Pictures.Skip(1).First().Id, Is.EqualTo(2));
                Assert.That(ws.Pictures.Skip(2).First().Id, Is.EqualTo(3));
                Assert.That(ws.Pictures.Skip(3).First().Id, Is.EqualTo(4));
            });
        }

        [Test]
        public void XLMarkerTests()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            XLMarker firstMarker = new XLMarker(ws.Cell(1, 10), new Point(100, 0));

            Assert.Multiple(() =>
            {
                Assert.That(firstMarker.ColumnNumber, Is.EqualTo(10));
                Assert.That(firstMarker.RowNumber, Is.EqualTo(1));
                Assert.That(firstMarker.Offset.X, Is.EqualTo(100));
                Assert.That(firstMarker.Offset.Y, Is.EqualTo(0));
            });
        }

        [Test]
        public void XLPictureTests()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png");
            var pic = ws.AddPicture(stream, XLPictureFormat.Png, "Image1")
                .WithPlacement(XLPicturePlacement.FreeFloating)
                .MoveTo(220, 155);

            Assert.Multiple(() =>
            {
                Assert.That(pic.Placement, Is.EqualTo(XLPicturePlacement.FreeFloating));
                Assert.That(pic.Name, Is.EqualTo("Image1"));
                Assert.That(pic.Format, Is.EqualTo(XLPictureFormat.Png));
                Assert.That(pic.OriginalWidth, Is.EqualTo(252));
                Assert.That(pic.OriginalHeight, Is.EqualTo(152));
                Assert.That(pic.Width, Is.EqualTo(252));
                Assert.That(pic.Height, Is.EqualTo(152));
                Assert.That(pic.Left, Is.EqualTo(220));
                Assert.That(pic.Top, Is.EqualTo(155));
            });
        }

        [Test]
        public void CanLoadFileWithImagesAndCopyImagesToNewSheet()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();
            Assert.That(ws.Pictures, Has.Count.EqualTo(2));

            var copy = ws.CopyTo("NewSheet");
            Assert.That(copy.Pictures, Has.Count.EqualTo(2));
        }

        [Test]
        public void CanDeletePictures()
        {
            using var ms = new MemoryStream();
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
                Assert.That(ws.Pictures, Has.Count.EqualTo(originalCount - 2));
            }
        }

        [Test]
        public void PictureRenameTests()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet("Images3");
            var picture = ws.Pictures.First();
            Assert.That(picture.Name, Is.EqualTo("Picture 1"));

            picture.Name = "picture 1";
            picture.Name = "pICture 1";
            picture.Name = "Picture 1";

            picture = ws.Pictures.Last();
            picture.Name = "new name";

            Assert.Throws<ArgumentException>(() => picture.Name = "Picture 1");
            Assert.Throws<ArgumentException>(() => picture.Name = "picTURE 1");
        }

        [Test]
        public void HandleDuplicatePictureIdsAcrossWorksheets()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet1");
            var ws2 = wb.AddWorksheet("Sheet2");

            using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png");
            (ws1 as XLWorksheet).AddPicture(stream, "Picture 1", 2);
            (ws1 as XLWorksheet).AddPicture(stream, "Picture 2", 3);

            //Internal method - used for loading files
            var pic = (ws2 as XLWorksheet).AddPicture(stream, "Picture 1", 2)
                .WithPlacement(XLPicturePlacement.FreeFloating)
                .MoveTo(220, 155) as XLPicture;

            var id = pic.Id;

            pic.Id = id;
            Assert.That(pic.Id, Is.EqualTo(id));

            pic.Id = 3;
            Assert.That(pic.Id, Is.EqualTo(3));

            pic.Id = id;

            var pic2 = (ws2 as XLWorksheet).AddPicture(stream, "Picture 2", 3)
                .WithPlacement(XLPicturePlacement.FreeFloating)
                .MoveTo(440, 300) as XLPicture;
        }

        [Test]
        public void CopyImageSameWorksheet()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");

            IXLPicture original;
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png"))
            {
                original = (ws1 as XLWorksheet).AddPicture(stream, "Picture 1", 2)
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .MoveTo(220, 155) as XLPicture;
            }

            var copy = original.Duplicate()
                .MoveTo(300, 200) as XLPicture;

            Assert.Multiple(() =>
            {
                Assert.That(ws1.Pictures, Has.Count.EqualTo(2));
                Assert.That(copy.Worksheet, Is.EqualTo(ws1));
                Assert.That(copy.Format, Is.EqualTo(original.Format));
                Assert.That(copy.Height, Is.EqualTo(original.Height));
                Assert.That(copy.Placement, Is.EqualTo(original.Placement));
                Assert.That(copy.TopLeftCell.ToString(), Is.EqualTo(original.TopLeftCell.ToString()));
                Assert.That(copy.Width, Is.EqualTo(original.Width));
                Assert.That(copy.ImageStream.ToArray(), Is.EqualTo(original.ImageStream.ToArray()), "Image streams differ");

                Assert.That(copy.Top, Is.EqualTo(200));
                Assert.That(copy.Left, Is.EqualTo(300));
            });
            Assert.Multiple(() =>
            {
                Assert.That(copy.Id, Is.Not.EqualTo(original.Id));
                Assert.That(copy.Name, Is.Not.EqualTo(original.Name));
            });
        }

        [Test]
        public void CopyImageDifferentWorksheets()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            IXLPicture original;
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png"))
            {
                original = (ws1 as XLWorksheet).AddPicture(stream, "Picture 1", 2)
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .MoveTo(220, 155) as XLPicture;
            }

            var ws2 = wb.Worksheets.Add("Sheet2");

            var copy = original.CopyTo(ws2);

            Assert.Multiple(() =>
            {
                Assert.That(ws1.Pictures, Has.Count.EqualTo(1));
                Assert.That(ws2.Pictures, Has.Count.EqualTo(1));

                Assert.That(copy.Worksheet, Is.EqualTo(ws2));

                Assert.That(copy.Format, Is.EqualTo(original.Format));
                Assert.That(copy.Height, Is.EqualTo(original.Height));
                Assert.That(copy.Left, Is.EqualTo(original.Left));
                Assert.That(copy.Name, Is.EqualTo(original.Name));
                Assert.That(copy.Placement, Is.EqualTo(original.Placement));
                Assert.That(copy.Top, Is.EqualTo(original.Top));
                Assert.That(copy.TopLeftCell.ToString(), Is.EqualTo(original.TopLeftCell.ToString()));
                Assert.That(copy.Width, Is.EqualTo(original.Width));
                Assert.That(copy.ImageStream.ToArray(), Is.EqualTo(original.ImageStream.ToArray()), "Image streams differ");
            });

            Assert.That(copy.Id, Is.Not.EqualTo(original.Id));
        }

        [Test]
        public void PictureShiftsWhenInsertingRows()
        {
            using var wb = new XLWorkbook();
            using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png");
            var ws = wb.Worksheets.Add("ImageShift");
            var picture = ws.AddPicture(stream, XLPictureFormat.Png, "PngImage")
                .MoveTo(ws.Cell(5, 2))
                .WithPlacement(XLPicturePlacement.Move);

            ws.Row(2).InsertRowsBelow(20);

            Assert.That(picture.TopLeftCell.Address.RowNumber, Is.EqualTo(25));
        }

        [Test]
        public void PictureNotFound()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Picture("dummy"));
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Pictures.Delete("dummy"));
        }

        [Test]
        public void CanCopyEmfPicture()
        {
            // #1621 - There are 2 Bmp Guids: ImageFormat.Bmp and ImageFormat.MemoryBmp
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Pictures\EmfPicture.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws1 = wb.Worksheets.First();
            var img1 = ws1.Pictures.First();

            var ws2 = wb.AddWorksheet();

            var img2 = img1.CopyTo(ws2);

            Assert.That(img2.Format, Is.EqualTo(XLPictureFormat.Emf));

            using var ms = new MemoryStream();
            wb.SaveAs(ms);

            ms.Seek(0, SeekOrigin.Begin);

            using var wb2 = new XLWorkbook(ms);
            ws2 = wb2.Worksheet("Sheet2");
            img2 = ws2.Pictures.First();
            Assert.That(img2.Format, Is.EqualTo(XLPictureFormat.Emf));
        }

        [Test]
        public void KeepOriginalDrawingShapesZOrder()
        {
            // File contains shapes and a picture in a mixed order.
            using var stream = TestHelper.GetStreamFromResource(@"Other.Pictures.ImageShapeZOrder-Input.xlsx");
            TestHelper.CreateAndCompare(
                () => new XLWorkbook(stream),
                @"Other\Pictures\ImageShapeZOrder-Output.xlsx");
        }
    }
}