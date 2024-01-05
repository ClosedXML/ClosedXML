using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Drawings
{
    [TestFixture]
    public class PictureTests
    {
        [TestCase("Other.Drawings.picture-webp.xlsx")]
        public void Can_load_and_save_workbook_with_image_type(string resourceWithImageType)
        {
            TestHelper.LoadSaveAndCompare(resourceWithImageType, resourceWithImageType);
        }
    }
}
