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
            using var wb = new XLWorkbook();
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

            Assert.Multiple(() =>
            {
                //Assert
                Assert.That(ws.Cell("A1").Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Cell("A2").Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
                Assert.That(ws.Cell("B1").Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Blue));
                Assert.That(ws.Cell("B2").Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Pink));

                Assert.That(ws.Cell("A1").Style.Font.FontName, Is.EqualTo("Arial"));
                Assert.That(ws.Cell("A2").Style.Font.FontName, Is.EqualTo("Times New Roman"));
                Assert.That(ws.Cell("B1").Style.Font.FontName, Is.EqualTo("Calibri"));
                Assert.That(ws.Cell("B2").Style.Font.FontName, Is.EqualTo("Cambria"));

                Assert.That(ws.Cell("A1").Style.Font.FontColor, Is.EqualTo(XLColor.PowderBlue));
                Assert.That(ws.Cell("A2").Style.Font.FontColor, Is.EqualTo(XLColor.PowderBlue));
                Assert.That(ws.Cell("B1").Style.Font.FontColor, Is.EqualTo(XLColor.PowderBlue));
                Assert.That(ws.Cell("B2").Style.Font.FontColor, Is.EqualTo(XLColor.PowderBlue));
            });
        }

        [Test]
        public void ChangeDetachedStyleAlignment()
        {
            var style = XLStyle.Default;

            style.Alignment.Horizontal = XLAlignmentHorizontalValues.Justify;

            Assert.That(style.Alignment.Horizontal, Is.EqualTo(XLAlignmentHorizontalValues.Justify));
        }

        [Test]
        public void ChangeDetachedStyleBorder()
        {
            var style = XLStyle.Default;

            style.Border.DiagonalBorder = XLBorderStyleValues.Double;

            Assert.That(style.Border.DiagonalBorder, Is.EqualTo(XLBorderStyleValues.Double));
        }

        [Test]
        public void ChangeDetachedStyleFill()
        {
            var style = XLStyle.Default;

            style.Fill.BackgroundColor = XLColor.Red;

            Assert.That(style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
        }

        [Test]
        public void ChangeDetachedStyleFont()
        {
            var style = XLStyle.Default;

            style.Font.FontSize = 50;

            Assert.That(style.Font.FontSize, Is.EqualTo(50));
        }

        [Test]
        public void ChangeDetachedStyleNumberFormat()
        {
            var style = XLStyle.Default;

            style.NumberFormat.Format = "YYYY";

            Assert.That(style.NumberFormat.Format, Is.EqualTo("YYYY"));
        }

        [Test]
        public void ChangeDetachedStyleProtection()
        {
            var style = XLStyle.Default;

            style.Protection.Hidden = true;

            Assert.That(style.Protection.Hidden, Is.True);
        }

        [Test]
        public void ChangeAttachedStyleAlignment()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var a1 = ws.Cell("A1");

            a1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Justify;

            Assert.That(a1.Style.Alignment.Horizontal, Is.EqualTo(XLAlignmentHorizontalValues.Justify));
        }
    }
}
