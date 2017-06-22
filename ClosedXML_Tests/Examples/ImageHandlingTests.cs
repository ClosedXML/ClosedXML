using ClosedXML_Examples;
using NUnit.Framework;

namespace ClosedXML_Tests.Examples
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
