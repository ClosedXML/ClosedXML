using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel.Misc
{
    /// <summary>
    ///     Summary description for UnitTest1
    /// </summary>
    [TestFixture]
    public class StylesTests
    {
        private static void SetupBorders(IXLRange range)
        {
            range.FirstRow().Cell(1).Style.Border.TopBorder = XLBorderStyleValues.None;
            range.FirstRow().Cell(2).Style.Border.TopBorder = XLBorderStyleValues.Thick;
            range.FirstRow().Cell(3).Style.Border.TopBorder = XLBorderStyleValues.Double;

            range.LastRow().Cell(1).Style.Border.BottomBorder = XLBorderStyleValues.None;
            range.LastRow().Cell(2).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
            range.LastRow().Cell(3).Style.Border.BottomBorder = XLBorderStyleValues.Double;

            range.FirstColumn().Cell(1).Style.Border.LeftBorder = XLBorderStyleValues.None;
            range.FirstColumn().Cell(2).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
            range.FirstColumn().Cell(3).Style.Border.LeftBorder = XLBorderStyleValues.Double;

            range.LastColumn().Cell(1).Style.Border.RightBorder = XLBorderStyleValues.None;
            range.LastColumn().Cell(2).Style.Border.RightBorder = XLBorderStyleValues.Thick;
            range.LastColumn().Cell(3).Style.Border.RightBorder = XLBorderStyleValues.Double;
        }

        [Test]
        public void InsideBorderTest()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLRange range = ws.Range("B2:D4");

            SetupBorders(range);

            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorderColor = XLColor.Red;

            IXLCell center = range.Cell(2, 2);

            Assert.AreEqual(XLColor.Red, center.Style.Border.TopBorderColor);
            Assert.AreEqual(XLColor.Red, center.Style.Border.BottomBorderColor);
            Assert.AreEqual(XLColor.Red, center.Style.Border.LeftBorderColor);
            Assert.AreEqual(XLColor.Red, center.Style.Border.RightBorderColor);

            Assert.AreEqual(XLBorderStyleValues.None, range.FirstRow().Cell(1).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thick, range.FirstRow().Cell(2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Double, range.FirstRow().Cell(3).Style.Border.TopBorder);

            Assert.AreEqual(XLBorderStyleValues.None, range.LastRow().Cell(1).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thick, range.LastRow().Cell(2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Double, range.LastRow().Cell(3).Style.Border.BottomBorder);

            Assert.AreEqual(XLBorderStyleValues.None, range.FirstColumn().Cell(1).Style.Border.LeftBorder);
            Assert.AreEqual(XLBorderStyleValues.Thick, range.FirstColumn().Cell(2).Style.Border.LeftBorder);
            Assert.AreEqual(XLBorderStyleValues.Double, range.FirstColumn().Cell(3).Style.Border.LeftBorder);

            Assert.AreEqual(XLBorderStyleValues.None, range.LastColumn().Cell(1).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thick, range.LastColumn().Cell(2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Double, range.LastColumn().Cell(3).Style.Border.RightBorder);
        }
    }
}