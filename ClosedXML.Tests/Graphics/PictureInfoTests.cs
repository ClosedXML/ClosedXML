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
        public void CanAddGif87Image()
        {
            AssertRasterImage("SampleImageGif87a.gif", XLPictureFormat.Gif, new Size(500, 200), 0, 0);
        }

        [Test]
        public void CanAddGif89Image()
        {
            AssertRasterImage("SampleImageGif89a.gif", XLPictureFormat.Gif, new Size(500, 200), 0, 0);
        }

        private static void AssertRasterImage(string imageName, XLPictureFormat expectedFormat, Size expectedPxSize, double expectedDpiX, double expectedDpiY)
        {
            using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream($"ClosedXML.Tests.Resource.Images.{imageName}");
            var info = DefaultGraphicEngine.Instance.Value.GetPictureInfo(stream, XLPictureFormat.Unknown);

            Assert.AreEqual(expectedFormat, info.Format);
            Assert.AreEqual(expectedPxSize, info.SizePx);
            Assert.AreEqual(expectedDpiX, info.DpiX);
            Assert.AreEqual(expectedDpiY, info.DpiY);
            Assert.AreEqual(Size.Empty, info.SizePhys);
        }
    }
}
