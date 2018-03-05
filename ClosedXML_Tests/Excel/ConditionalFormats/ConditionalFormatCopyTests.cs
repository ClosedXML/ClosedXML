using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel.ConditionalFormats
{
    [TestFixture]
    public class ConditionalFormatCopyTests
    {
        [Test]
        public void StylesAreCreatedDuringCopy()
        {
            var wb = new XLWorkbook();
            XLWorksheet ws = (XLWorksheet)wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("A1:A1").AddConditionalFormat());
            var wb2 = new XLWorkbook();
            XLWorksheet ws2 = (XLWorksheet)wb2.Worksheets.Add("Sheet2");
            ws2.FirstCell().CopyFrom(ws.FirstCell());
            Assert.That(ws2.ConditionalFormats.First().Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Blue)); //Added blue style

        }

        private static void SetFormat1(IXLConditionalFormat format)
        {
            format.WhenEquals("=" + format.Range.FirstCell().CellRight(4).Address.ToStringRelative()).Fill.SetBackgroundColor(XLColor.Blue);
        }
    }
}
