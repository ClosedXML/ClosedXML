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
    }
}