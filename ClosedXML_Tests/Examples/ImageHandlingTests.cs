using ClosedXML_Examples;
using NUnit.Framework;

namespace ClosedXML_Tests.Examples
{
    [TestFixture]
    public class ImageHandlingTests
    {
        [Test]
        public void ImageHandling()
        {
            TestHelper.RunTestExample<ImageAnchors>(@"ImageHandling\ImageHandling.xlsx");
        }

        [Test]
        public void ImageFormats()
        {
            TestHelper.RunTestExample<ImageFormats>(@"ImageHandling\ImageFormats.xlsx");
        }
    }
}