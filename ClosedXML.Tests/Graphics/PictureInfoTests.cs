using ClosedXML.Excel.Drawings;
using ClosedXML.Graphics;
using NUnit.Framework;
using System.Drawing;
using System.Reflection;

namespace ClosedXML.Tests.Graphics
{
    [TestFixture]
    public class PictureInfoTests
    {
        [Test]
        public void CanReadGif87Image()
        {
            AssertRasterImage("SampleImageGif87a.gif", XLPictureFormat.Gif, new Size(500, 200), 0, 0);
        }

        [Test]
        public void CanReadGif89Image()
        {
            AssertRasterImage("SampleImageGif89a.gif", XLPictureFormat.Gif, new Size(500, 200), 0, 0);
        }

        [TestCase("SampleImageBmpWin24bit.bmp")]
        [TestCase("SampleImageBmpWin8bit.bmp")]
        [TestCase("SampleImageBmpWin4bit.bmp")]
        [TestCase("SampleImageBmpWin24bit.bmp")]
        public void CanReadBmpImageV3AndFurther(string imageName)
        {
            AssertRasterImage(imageName, XLPictureFormat.Bmp, new Size(167, 51), 80.645d, 80.645d);
        }

        [Test]
        public void CanReadBmpV1()
        {
            AssertRasterImage("SampleImageBmpV1.bmp", XLPictureFormat.Bmp, new Size(150, 50), 0, 0);
        }

        [Test]
        public void CanReadTiffWithBigEndianEncoding()
        {
            AssertRasterImage("SampleImageTiffBigEndian.tiff", XLPictureFormat.Tiff, new Size(130, 45), 96, 96);
        }

        [Test]
        public void CanReadTiffWithLittleEndianEncoding()
        {
            AssertRasterImage("SampleImageTiffLittleEndian.tiff", XLPictureFormat.Tiff, new Size(130, 45), 96, 96);
        }

        private static void AssertRasterImage(string imageName, XLPictureFormat expectedFormat, Size expectedPxSize, double expectedDpiX, double expectedDpiY)
        {
            using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream($"ClosedXML.Tests.Resource.Images.{imageName}");
            var info = DefaultGraphicEngine.Instance.Value.GetPictureInfo(stream, XLPictureFormat.Unknown);

            Assert.AreEqual(expectedFormat, info.Format);
            Assert.AreEqual(expectedPxSize, info.SizePx);
            Assert.AreEqual(expectedDpiX, info.DpiX, 0.01);
            Assert.AreEqual(expectedDpiY, info.DpiY, 0.01);
            Assert.AreEqual(Size.Empty, info.SizePhys);
        }
    }
}
