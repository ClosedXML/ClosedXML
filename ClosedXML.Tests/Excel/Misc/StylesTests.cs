using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Misc
{
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

        [Test]
        public void ResolveThemeColors()
        {
            using (var wb = new XLWorkbook())
            {
                string color;
                color = wb.Theme.ResolveThemeColor(XLThemeColor.Accent1).Color.ToHex();
                Assert.AreEqual("FF4F81BD", color);

                color = wb.Theme.ResolveThemeColor(XLThemeColor.Background1).Color.ToHex();
                Assert.AreEqual("FFFFFFFF", color);
            }
        }

        [Theory]
        public void CanResolveAllThemeColors(XLThemeColor themeColor)
        {
            var theme = new XLWorkbook().Theme;
            var color = theme.ResolveThemeColor(themeColor);
            Assert.IsNotNull(color);
        }

        [Test]
        public void SetStyleViaRowReference()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Style
                   .Font.SetFontSize(8)
                   .Font.SetFontColor(XLColor.Green)
                   .Font.SetBold(true);

                var row = ws.Row(1);
                ws.Cell(1, 1).Value = "Test";
                row.Cell(2).Value = "Test";
                row.Cells(3, 3).Value = "Test";

                foreach (var cell in ws.CellsUsed())
                {
                    Assert.AreEqual(8, ws.Cell("A1").Style.Font.FontSize);
                    Assert.AreEqual(XLColor.Green, ws.Cell("B1").Style.Font.FontColor);
                    Assert.AreEqual(true, ws.Cell("C1").Style.Font.Bold);
                }
            }
        }
    }
}
