using NUnit.Framework;

namespace ClosedXML.Tests.Excel.PageSetup
{
    [TestFixture]
    public class PageLayoutTests
    {
        [Test]
        public void FirstPageNumber_can_be_negative()
        {
            TestHelper.CreateSaveLoadAssert(
                (_, ws) => ws.PageSetup.FirstPageNumber = -3,
                (_, ws) => Assert.That(ws.PageSetup.FirstPageNumber, Is.EqualTo(-3)),
                @"Other\PageSetup\Negative_first_page_number.xlsx");
        }
    }
}
