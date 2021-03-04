using ClosedXML.Excel;
using NUnit.Framework;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class HeaderFooterTests
    {
        [Test]
        public void CanChangeWorksheetHeader()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            ws.PageSetup.Header.Center.AddText("Initial page header", XLHFOccurrence.EvenPages);

            var ms = new MemoryStream();
            wb.SaveAs(ms, true);

            wb = new XLWorkbook(ms);
            ws = wb.Worksheets.First();

            ws.PageSetup.Header.Center.Clear();
            ws.PageSetup.Header.Center.AddText("Changed header", XLHFOccurrence.EvenPages);

            wb.SaveAs(ms, true);

            wb = new XLWorkbook(ms);
            ws = wb.Worksheets.First();

            var newHeader = ws.PageSetup.Header.Center.GetText(XLHFOccurrence.EvenPages);
            Assert.AreEqual("Changed header", newHeader);
        }

        [TestCase("")]
        [TestCase("&L&C&\"Arial\"&9 19-10-2017 \n&9&\"Arial\" &P    &N &R")] // https://github.com/ClosedXML/ClosedXML/issues/563
        public void CanSetHeaderFooter(string s)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                {
                    var header = ws.PageSetup.Header as XLHeaderFooter;
                    header.SetInnerText(XLHFOccurrence.AllPages, s);
                }
            }
        }
    }
}
