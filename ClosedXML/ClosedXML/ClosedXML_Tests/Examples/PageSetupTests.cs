using ClosedXML_Examples.PageSetup;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Examples
{
    [TestClass]
    public class PageSetupTests
    {
        [TestMethod]
        public void HeaderFooters()
        {
            TestHelper.RunTestExample<HeaderFooters>(@"PageSetup\HeaderFooters.xlsx");
        }
        [TestMethod]
        public void Margins()
        {
            TestHelper.RunTestExample<Margins>(@"PageSetup\Margins.xlsx");
        }
        [TestMethod]
        public void Page()
        {
            TestHelper.RunTestExample<Page>(@"PageSetup\Page.xlsx");
        }
        [TestMethod]
        public void Sheets()
        {
            TestHelper.RunTestExample<Sheets>(@"PageSetup\Sheets.xlsx");
        }
        [TestMethod]
        public void SheetTab()
        {
            TestHelper.RunTestExample<SheetTab>(@"PageSetup\SheetTab.xlsx");
        }
        [TestMethod]
        public void TwoPages()
        {
            TestHelper.RunTestExample<TwoPages>(@"PageSetup\TwoPages.xlsx");
        }

    }
}