using ClosedXML.Excel.Drawings;
using ClosedXML.Graphics;
using NUnit.Framework;
using System.Drawing;
using System.Reflection;

namespace ClosedXML.SixLabors.Tests
{
    public class ImageTests
    {
        [Test]
        public void CanReadEmf()
        {
            var engine = new SixLaborsEngine();
            using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.SixLabors.Tests.Resource.ReadImage.emf");
            var picture = engine.GetPictureMetadata(stream, XLPictureFormat.Unknown);
            Assert.That(picture.Format, Is.EqualTo(XLPictureFormat.Emf));
            Assert.That(picture.SizePx, Is.EqualTo(new Size(924, 927)));
            Assert.That(picture.SizePhys, Is.EqualTo(new Size(28844, 28938)));
        }
    }
}
