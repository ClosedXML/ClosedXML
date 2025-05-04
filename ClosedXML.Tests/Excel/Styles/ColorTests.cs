using ClosedXML.Excel;
using ClosedXML.Utils;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using System.Globalization;
using System.Threading;
using Color = System.Drawing.Color;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class ColorTests
    {
        [Test]
        public void ColorEqualOperatorInPlace()
        {
            Assert.That(XLColor.Black == XLColor.Black, Is.True);
        }

        [Test]
        public void ColorNotEqualOperatorInPlace()
        {
            Assert.That(XLColor.Black != XLColor.Black, Is.False);
        }

        [Test]
        public void ColorNamedVsHTML()
        {
            Assert.That(XLColor.Black == XLColor.FromHtml("#000000"), Is.True);
        }

        [Test]
        public void DefaultColorIndex64isTransparentWhite()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            XLColor color = ws.FirstCell().Style.Fill.BackgroundColor;
            Assert.Multiple(() =>
            {
                Assert.That(color.ColorType, Is.EqualTo(XLColorType.Indexed));
                Assert.That(color.Indexed, Is.EqualTo(64));
                Assert.That(color.Color, Is.EqualTo(Color.Transparent));
            });
        }

        [Test]
        public void CanConvertXLColorToColorType()
        {
            var xlColor1 = XLColor.Red;
            var xlColor2 = XLColor.FromIndex(20);
            var xlColor3 = XLColor.FromTheme(XLThemeColor.Accent1);
            var xlColor4 = XLColor.FromTheme(XLThemeColor.Accent2, 0.4);

            var color1 = new ForegroundColor().FromClosedXMLColor<ForegroundColor>(xlColor1);
            var color2 = new ForegroundColor().FromClosedXMLColor<ForegroundColor>(xlColor2);
            var color3 = new BackgroundColor().FromClosedXMLColor<BackgroundColor>(xlColor3);
            var color4 = new BackgroundColor().FromClosedXMLColor<BackgroundColor>(xlColor4);

            Assert.That(color1.Rgb.Value, Is.EqualTo("FFFF0000"));
            Assert.That(color1.Indexed, Is.Null);
            Assert.That(color1.Theme, Is.Null);
            Assert.That(color1.Tint, Is.Null);

            Assert.That(color2.Rgb, Is.Null);
            Assert.That(color2.Indexed.Value, Is.EqualTo(20));
            Assert.That(color2.Theme, Is.Null);
            Assert.That(color2.Tint, Is.Null);

            Assert.That(color3.Rgb, Is.Null);
            Assert.That(color3.Indexed, Is.Null);
            Assert.That(color3.Theme.Value, Is.EqualTo(4));
            Assert.That(color3.Tint, Is.Null);

            Assert.That(color4.Rgb, Is.Null);
            Assert.That(color4.Indexed, Is.Null);
            Assert.Multiple(() =>
            {
                Assert.That(color4.Theme.Value, Is.EqualTo(5));
                Assert.That(color4.Tint.Value, Is.EqualTo(0.4));
            });
        }

        [Test]
        public void CanConvertXlColorToX14ColorType()
        {
            var xlColor1 = XLColor.Red;
            var xlColor2 = XLColor.FromIndex(20);
            var xlColor3 = XLColor.FromTheme(XLThemeColor.Accent1);
            var xlColor4 = XLColor.FromTheme(XLThemeColor.Accent2, 0.4);

            var color1 = new X14.AxisColor().FromClosedXMLColor<X14.AxisColor>(xlColor1);
            var color2 = new X14.BorderColor().FromClosedXMLColor<X14.BorderColor>(xlColor2);
            var color3 = new X14.FillColor().FromClosedXMLColor<X14.FillColor>(xlColor3);
            var color4 = new X14.HighMarkerColor().FromClosedXMLColor<X14.HighMarkerColor>(xlColor4);

            Assert.That(color1.Rgb.Value, Is.EqualTo("FFFF0000"));
            Assert.That(color1.Indexed, Is.Null);
            Assert.That(color1.Theme, Is.Null);
            Assert.That(color1.Tint, Is.Null);

            Assert.That(color2.Rgb, Is.Null);
            Assert.That(color2.Indexed.Value, Is.EqualTo(20));
            Assert.That(color2.Theme, Is.Null);
            Assert.That(color2.Tint, Is.Null);

            Assert.That(color3.Rgb, Is.Null);
            Assert.That(color3.Indexed, Is.Null);
            Assert.That(color3.Theme.Value, Is.EqualTo(4));
            Assert.That(color3.Tint, Is.Null);

            Assert.That(color4.Rgb, Is.Null);
            Assert.That(color4.Indexed, Is.Null);
            Assert.Multiple(() =>
            {
                Assert.That(color4.Theme.Value, Is.EqualTo(5));
                Assert.That(color4.Tint.Value, Is.EqualTo(0.4));
            });
        }

        [Test]
        public void CanConvertColorTypeToXlColor()
        {
            var color1 = new ForegroundColor { Rgb = new DocumentFormat.OpenXml.HexBinaryValue("FFFF0000") };
            var color2 = new ForegroundColor { Indexed = new DocumentFormat.OpenXml.UInt32Value((uint)20) };
            var color3 = new BackgroundColor { Theme = new DocumentFormat.OpenXml.UInt32Value((uint)4) };
            var color4 = new BackgroundColor
            {
                Theme = new DocumentFormat.OpenXml.UInt32Value((uint)4),
                Tint = new DocumentFormat.OpenXml.DoubleValue(0.4)
            };

            var xlColor1 = color1.ToClosedXMLColor();
            var xlColor2 = color2.ToClosedXMLColor();
            var xlColor3 = color3.ToClosedXMLColor();
            var xlColor4 = color4.ToClosedXMLColor();

            Assert.Multiple(() =>
            {
                Assert.That(xlColor1.ColorType, Is.EqualTo(XLColorType.Color));
                Assert.That(xlColor1.Color, Is.EqualTo(XLColor.Red.Color));

                Assert.That(xlColor2.ColorType, Is.EqualTo(XLColorType.Indexed));
                Assert.That(xlColor2.Indexed, Is.EqualTo(20));

                Assert.That(xlColor3.ColorType, Is.EqualTo(XLColorType.Theme));
                Assert.That(xlColor3.ThemeColor, Is.EqualTo(XLThemeColor.Accent1));
                Assert.That(xlColor3.ThemeTint, Is.EqualTo(0).Within(XLHelper.Epsilon));

                Assert.That(xlColor4.ColorType, Is.EqualTo(XLColorType.Theme));
                Assert.That(xlColor4.ThemeColor, Is.EqualTo(XLThemeColor.Accent1));
                Assert.That(xlColor4.ThemeTint, Is.EqualTo(0.4).Within(XLHelper.Epsilon));
            });
        }

        [Test]
        public void CanConvertX14ColorTypeToXlColor()
        {
            var color1 = new X14.AxisColor { Rgb = new DocumentFormat.OpenXml.HexBinaryValue("FFFF0000") };
            var color2 = new X14.BorderColor { Indexed = new DocumentFormat.OpenXml.UInt32Value((uint)20) };
            var color3 = new X14.FillColor { Theme = new DocumentFormat.OpenXml.UInt32Value((uint)4) };
            var color4 = new X14.HighMarkerColor
            {
                Theme = new DocumentFormat.OpenXml.UInt32Value((uint)4),
                Tint = new DocumentFormat.OpenXml.DoubleValue(0.4)
            };

            var xlColor1 = color1.ToClosedXMLColor();
            var xlColor2 = color2.ToClosedXMLColor();
            var xlColor3 = color3.ToClosedXMLColor();
            var xlColor4 = color4.ToClosedXMLColor();

            Assert.Multiple(() =>
            {
                Assert.That(xlColor1.ColorType, Is.EqualTo(XLColorType.Color));
                Assert.That(xlColor1.Color, Is.EqualTo(XLColor.Red.Color));

                Assert.That(xlColor2.ColorType, Is.EqualTo(XLColorType.Indexed));
                Assert.That(xlColor2.Indexed, Is.EqualTo(20));

                Assert.That(xlColor3.ColorType, Is.EqualTo(XLColorType.Theme));
                Assert.That(xlColor3.ThemeColor, Is.EqualTo(XLThemeColor.Accent1));
                Assert.That(xlColor3.ThemeTint, Is.EqualTo(0).Within(XLHelper.Epsilon));

                Assert.That(xlColor4.ColorType, Is.EqualTo(XLColorType.Theme));
                Assert.That(xlColor4.ThemeColor, Is.EqualTo(XLThemeColor.Accent1));
                Assert.That(xlColor4.ThemeTint, Is.EqualTo(0.4).Within(XLHelper.Epsilon));
            });
        }

        [Test]
        public void CanParseColorWithHashAsCultureLineSeparator()
        {
            // https://github.com/ClosedXML/ClosedXML/issues/675
            var culture = CultureInfo.CreateSpecificCulture("en-US");
            culture.TextInfo.ListSeparator = "#";
            Thread.CurrentThread.CurrentCulture = culture;
            var color = XLColor.FromHtml("#FF008000");
            Assert.That(color, Is.EqualTo(XLColor.Green));
        }
    }
}
