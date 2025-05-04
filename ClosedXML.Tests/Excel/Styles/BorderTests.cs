using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Styles
{
    public class BorderTests
    {
        [Test]
        public void SetInsideBorderPreservesOutsideBorders()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cells("B2:C2").Style
                .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                .Border.SetOutsideBorderColor(XLColor.FromTheme(XLThemeColor.Accent1, 0.5));

            Assert.Multiple(() =>
            {
                //Check pre-conditions
                Assert.That(ws.Cell("B2").Style.Border.LeftBorder, Is.EqualTo(XLBorderStyleValues.Thin));
                Assert.That(ws.Cell("B2").Style.Border.RightBorder, Is.EqualTo(XLBorderStyleValues.Thin));
                Assert.That(ws.Cell("B2").Style.Border.LeftBorderColor.ThemeColor, Is.EqualTo(XLThemeColor.Accent1));
                Assert.That(ws.Cell("B2").Style.Border.RightBorderColor.ThemeColor, Is.EqualTo(XLThemeColor.Accent1));
            });

            ws.Range("B2:C2").Style.Border.SetInsideBorder(XLBorderStyleValues.None);

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("B2").Style.Border.LeftBorder, Is.EqualTo(XLBorderStyleValues.Thin));
                Assert.That(ws.Cell("B2").Style.Border.RightBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("C2").Style.Border.LeftBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("C2").Style.Border.RightBorder, Is.EqualTo(XLBorderStyleValues.Thin));
                Assert.That(ws.Cell("B2").Style.Border.LeftBorderColor.ThemeColor, Is.EqualTo(XLThemeColor.Accent1));
                Assert.That(ws.Cell("C2").Style.Border.RightBorderColor.ThemeColor, Is.EqualTo(XLThemeColor.Accent1));
            });
        }
    }
}
