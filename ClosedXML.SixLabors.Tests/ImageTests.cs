using ClosedXML.Excel.Drawings;
using ClosedXML.Graphics;
using NUnit.Framework;
using System.Drawing;
using System.Reflection;

namespace ClosedXML.SixLabors.Tests
{
    [TestFixture]
    public class ReadMetadataTests
    {
        private readonly IXLGraphicEngine _engine = new SixLaborsEngine();

        [Test]
        public void CanReadPng()
        {
            var picture = GetMetadata("SamplePng.png");
            Assert.That(picture.Format, Is.EqualTo(XLPictureFormat.Png));
            Assert.That(picture.SizePx, Is.EqualTo(new Size(252, 152)));
            Assert.That(picture.SizePhys, Is.EqualTo(Size.Empty));
        }

        [Test]
        public void CanReadJfif()
        {
            var picture = GetMetadata("SampleJfif.jpg");
            Assert.That(picture.Format, Is.EqualTo(XLPictureFormat.Jpeg));
            Assert.That(picture.SizePx, Is.EqualTo(new Size(176, 270)));
            Assert.That(picture.SizePhys, Is.EqualTo(Size.Empty));
        }

        [Test]
        public void CanReadExif()
        {
            var picture = GetMetadata("SampleExif.jpg");
            Assert.That(picture.Format, Is.EqualTo(XLPictureFormat.Jpeg));
            Assert.That(picture.SizePx, Is.EqualTo(new Size(252, 152)));
            Assert.That(picture.SizePhys, Is.EqualTo(Size.Empty));
        }

        [Test]
        public void CanReadGif()
        {
            var picture = GetMetadata("SampleGif.gif");
            Assert.That(picture.Format, Is.EqualTo(XLPictureFormat.Gif));
            Assert.That(picture.SizePx, Is.EqualTo(new Size(250, 210)));
            Assert.That(picture.SizePhys, Is.EqualTo(Size.Empty));
        }

        [Test]
        public void CanReadBmp()
        {
            var picture = GetMetadata("SampleBmp.bmp");
            Assert.That(picture.Format, Is.EqualTo(XLPictureFormat.Bmp));
            Assert.That(picture.SizePx, Is.EqualTo(new Size(247, 89)));
            Assert.That(picture.SizePhys, Is.EqualTo(Size.Empty));
        }

        [Test]
        public void CanReadTiff()
        {
            var picture = GetMetadata("SampleTiff.tif");
            Assert.That(picture.Format, Is.EqualTo(XLPictureFormat.Tiff));
            Assert.That(picture.SizePx, Is.EqualTo(new Size(212, 146)));
            Assert.That(picture.SizePhys, Is.EqualTo(Size.Empty));
        }

        [Test]
        public void CanReadEmf()
        {
            var picture = GetMetadata("SampleEmf.emf");
            Assert.That(picture.Format, Is.EqualTo(XLPictureFormat.Emf));
            Assert.That(picture.SizePx, Is.EqualTo(new Size(924, 927)));
            Assert.That(picture.SizePhys, Is.EqualTo(new Size(28844, 28938)));
        }

        private XLPictureInfo GetMetadata(string resourceImage)
        {
            using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream($"ClosedXML.SixLabors.Tests.Resource.{resourceImage}");
            return _engine.GetPictureInfo(stream, XLPictureFormat.Unknown);
        }
    }
}
