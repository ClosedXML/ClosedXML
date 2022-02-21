using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Styles
{
    public class BorderTests
    {
        [Test]
        public void SetInsideBorderPreservesOutsideBorders()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();

                ws.Cells("B2:C2").Style
                    .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                    .Border.SetOutsideBorderColor(XLColor.FromTheme(XLThemeColor.Accent1, 0.5));

                //Check pre-conditions
                Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell("B2").Style.Border.LeftBorder);
                Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell("B2").Style.Border.RightBorder);
                Assert.AreEqual(XLThemeColor.Accent1, ws.Cell("B2").Style.Border.LeftBorderColor.ThemeColor);
                Assert.AreEqual(XLThemeColor.Accent1, ws.Cell("B2").Style.Border.RightBorderColor.ThemeColor);

                ws.Range("B2:C2").Style.Border.SetInsideBorder(XLBorderStyleValues.None);

                Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell("B2").Style.Border.LeftBorder);
                Assert.AreEqual(XLBorderStyleValues.None, ws.Cell("B2").Style.Border.RightBorder);
                Assert.AreEqual(XLBorderStyleValues.None, ws.Cell("C2").Style.Border.LeftBorder);
                Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell("C2").Style.Border.RightBorder);
                Assert.AreEqual(XLThemeColor.Accent1, ws.Cell("B2").Style.Border.LeftBorderColor.ThemeColor);
                Assert.AreEqual(XLThemeColor.Accent1, ws.Cell("C2").Style.Border.RightBorderColor.ThemeColor);
            }
        }
    }
}
