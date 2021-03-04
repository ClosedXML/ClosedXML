using ClosedXML.Examples;
using NUnit.Framework;

namespace ClosedXML.Tests.Examples
{
    [TestFixture]
    public class ImageHandlingTests
    {
        [Test]
        public void ImageAnchors()
        {
            TestHelper.RunTestExample<ImageAnchors>(@"ImageHandling\ImageAnchors.xlsx");
        }

        [Test]
        public void ImageFormats()
        {
            TestHelper.RunTestExample<ImageFormats>(@"ImageHandling\ImageFormats.xlsx");
        }
    }
}
