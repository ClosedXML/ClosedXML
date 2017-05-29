using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using NUnit.Framework;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class PictureTests
    {
        [Test]
        public void XLMarkerTests()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            XLMarker firstMarker = new XLMarker
            {
                ColumnId = 1,
                RowId = 1,
                ColumnOffset = 100,
                RowOffset = 0
            };

            firstMarker.ColumnId = 10;

            Assert.AreEqual(10, firstMarker.ColumnId);
            Assert.AreEqual(1, firstMarker.RowId);
            Assert.AreEqual(100, firstMarker.ColumnOffset);
            Assert.AreEqual(0, firstMarker.RowOffset);
            //Assert.AreEqual(9, firstMarker.GetZeroBasedColumn());
            //Assert.AreEqual(0, firstMarker.GetZeroBasedRow());

            Assert.Throws(typeof(ArgumentOutOfRangeException), delegate { firstMarker.RowId = 0; });
            Assert.Throws(typeof(ArgumentOutOfRangeException), delegate { firstMarker.ColumnId = 0; });
            Assert.Throws(typeof(ArgumentOutOfRangeException),
                            delegate { firstMarker.RowId = XLHelper.MaxRowNumber + 1; });
            Assert.Throws(typeof(ArgumentOutOfRangeException),
                            delegate { firstMarker.ColumnId = XLHelper.MaxColumnNumber + 1; });
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
                        .AtPosition(220, 155);

                    Assert.AreEqual(XLPicturePlacement.FreeFloating, pic.Placement);
                    Assert.AreEqual("Image1", pic.Name);
                    Assert.AreEqual(XLPictureFormat.Png, pic.Format);
                    Assert.AreEqual(252, pic.OriginalWidth);
                    Assert.AreEqual(152, pic.OriginalHeight);
                    Assert.AreEqual(252, pic.Width);
                    Assert.AreEqual(152, pic.Height);
                    Assert.AreEqual(220, pic.Left);
                    Assert.AreEqual(155, pic.Top);
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
        public void CanAddPictureFromFile()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                var path = Path.ChangeExtension(Path.GetTempFileName(), "png");

                using (var resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Tests.Resource.Images.SampleImage.jpg"))
                using (var fileStream = File.Create(path))
                {
                    resourceStream.Seek(0, SeekOrigin.Begin);
                    resourceStream.CopyTo(fileStream);
                    fileStream.Close();
                }

                var picture = ws.AddPicture(path)
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .AtPosition(50, 50);

                Assert.AreEqual(XLPictureFormat.Jpeg, picture.Format);
                Assert.AreEqual(1365, picture.Width);
                Assert.AreEqual(1365, picture.Height);
            }
        }

        [Test]
        public void CanAddPictureFromBitmap()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                using (var resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Tests.Resource.Images.SampleImage.jpg"))
                using (var bitmap = Bitmap.FromStream(resourceStream) as Bitmap)
                {
                    var picture = ws.AddPicture(bitmap, "MyPicture")
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .AtPosition(50, 50)
                        .WithSize(200, 200);

                    Assert.AreEqual(XLPictureFormat.Jpeg, picture.Format);
                    Assert.AreEqual(200, picture.Width);
                    Assert.AreEqual(200, picture.Height);
                }
            }
        }

        [Test]
        public void CanScaleImage()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                using (var resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Tests.Resource.Images.ImageHandling.png"))
                using (var bitmap = Bitmap.FromStream(resourceStream) as Bitmap)
                {
                    var pic = ws.AddPicture(bitmap, "MyPicture")
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .AtPosition(50, 50);

                    Assert.AreEqual(252, pic.OriginalWidth);
                    Assert.AreEqual(152, pic.OriginalHeight);
                    Assert.AreEqual(252, pic.Width);
                    Assert.AreEqual(152, pic.Height);

                    pic.ScaleHeight(0.7);
                    pic.ScaleWidth(1.2);

                    Assert.AreEqual(252, pic.OriginalWidth);
                    Assert.AreEqual(152, pic.OriginalHeight);
                    Assert.AreEqual(302, pic.Width);
                    Assert.AreEqual(106, pic.Height);

                    pic.ScaleHeight(0.7);
                    pic.ScaleWidth(1.2);

                    Assert.AreEqual(252, pic.OriginalWidth);
                    Assert.AreEqual(152, pic.OriginalHeight);
                    Assert.AreEqual(362, pic.Width);
                    Assert.AreEqual(74, pic.Height);

                    pic.ScaleHeight(0.8, true);
                    pic.ScaleWidth(1.1, true);

                    Assert.AreEqual(252, pic.OriginalWidth);
                    Assert.AreEqual(152, pic.OriginalHeight);
                    Assert.AreEqual(277, pic.Width);
                    Assert.AreEqual(122, pic.Height);
                }
            }
        }
    }
}
