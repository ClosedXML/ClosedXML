using ClosedXML.Examples.PageSetup;
using NUnit.Framework;

namespace ClosedXML.Tests.Examples
{
    [TestFixture]
    public class PageSetupTests
    {
        [Test]
        public void HeaderFooters()
        {
            TestHelper.RunTestExample<HeaderFooters>(@"PageSetup\HeaderFooters.xlsx");
        }

        [Test]
        public void Margins()
        {
            TestHelper.RunTestExample<Margins>(@"PageSetup\Margins.xlsx");
        }

        [Test]
        public void Page()
        {
            TestHelper.RunTestExample<Page>(@"PageSetup\Page.xlsx");
        }

        [Test]
        public void SheetTab()
        {
            TestHelper.RunTestExample<SheetTab>(@"PageSetup\SheetTab.xlsx");
        }

        [Test]
        public void Sheets()
        {
            TestHelper.RunTestExample<Sheets>(@"PageSetup\Sheets.xlsx");
        }

        [Test]
        public void TwoPages()
        {
            TestHelper.RunTestExample<TwoPages>(@"PageSetup\TwoPages.xlsx");
        }
    }
}