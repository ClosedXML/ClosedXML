using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Styles
{
    [TestFixture]
    public class StyleChangeTests
    {
        [Test]
        public void ChangeFontColorDoesNotAffectOtherProperties()
        {
            using (var wb = new XLWorkbook())
            {
                // Arrange
                var ws = wb.AddWorksheet("Sheet1");
                var a1 = ws.Cell("A1");
                var a2 = ws.Cell("A2");
                var b1 = ws.Cell("B1");
                var b2 = ws.Cell("B2");

                ws.Range("A1:B2").Value = "Test";

                a1.Style.Fill.BackgroundColor = XLColor.Red;
                a2.Style.Fill.BackgroundColor = XLColor.Green;
                b1.Style.Fill.BackgroundColor = XLColor.Blue;
                b2.Style.Fill.BackgroundColor = XLColor.Pink;

                a1.Style.Font.FontName = "Arial";
                a2.Style.Font.FontName = "Times New Roman";
                b1.Style.Font.FontName = "Calibri";
                b2.Style.Font.FontName = "Cambria";

                // Act
                ws.Range("A1:B2").Style.Font.FontColor = XLColor.PowderBlue;

                //Assert
                Assert.AreEqual(XLColor.Red, ws.Cell("A1").Style.Fill.BackgroundColor);
                Assert.AreEqual(XLColor.Green, ws.Cell("A2").Style.Fill.BackgroundColor);
                Assert.AreEqual(XLColor.Blue, ws.Cell("B1").Style.Fill.BackgroundColor);
                Assert.AreEqual(XLColor.Pink, ws.Cell("B2").Style.Fill.BackgroundColor);

                Assert.AreEqual("Arial", ws.Cell("A1").Style.Font.FontName);
                Assert.AreEqual("Times New Roman", ws.Cell("A2").Style.Font.FontName);
                Assert.AreEqual("Calibri", ws.Cell("B1").Style.Font.FontName);
                Assert.AreEqual("Cambria", ws.Cell("B2").Style.Font.FontName);

                Assert.AreEqual(XLColor.PowderBlue, ws.Cell("A1").Style.Font.FontColor);
                Assert.AreEqual(XLColor.PowderBlue, ws.Cell("A2").Style.Font.FontColor);
                Assert.AreEqual(XLColor.PowderBlue, ws.Cell("B1").Style.Font.FontColor);
                Assert.AreEqual(XLColor.PowderBlue, ws.Cell("B2").Style.Font.FontColor);
            }
        }

        [Test]
        public void ChangeDetachedStyleAlignment()
        {
            var style = XLStyle.Default;

            style.Alignment.Horizontal = XLAlignmentHorizontalValues.Justify;

            Assert.AreEqual(XLAlignmentHorizontalValues.Justify, style.Alignment.Horizontal);
        }

        [Test]
        public void ChangeDetachedStyleBorder()
        {
            var style = XLStyle.Default;

            style.Border.DiagonalBorder = XLBorderStyleValues.Double;

            Assert.AreEqual(XLBorderStyleValues.Double, style.Border.DiagonalBorder);
        }

        [Test]
        public void ChangeDetachedStyleFill()
        {
            var style = XLStyle.Default;

            style.Fill.BackgroundColor = XLColor.Red;

            Assert.AreEqual(XLColor.Red, style.Fill.BackgroundColor);
        }

        [Test]
        public void ChangeDetachedStyleFont()
        {
            var style = XLStyle.Default;

            style.Font.FontSize = 50;

            Assert.AreEqual(50, style.Font.FontSize);
        }

        [Test]
        public void ChangeDetachedStyleNumberFormat()
        {
            var style = XLStyle.Default;

            style.NumberFormat.Format = "YYYY";

            Assert.AreEqual("YYYY", style.NumberFormat.Format);
        }

        [Test]
        public void ChangeDetachedStyleProtection()
        {
            var style = XLStyle.Default;

            style.Protection.Hidden = true;

            Assert.AreEqual(true, style.Protection.Hidden);
        }

        [Test]
        public void ChangeAttachedStyleAlignment()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                var a1 = ws.Cell("A1");

                a1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Justify;

                Assert.AreEqual(XLAlignmentHorizontalValues.Justify, a1.Style.Alignment.Horizontal);
            }
        }
    }
}
