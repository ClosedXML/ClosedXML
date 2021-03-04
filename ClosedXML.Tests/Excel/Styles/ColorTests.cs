using ClosedXML.Excel;
using ClosedXML.Utils;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
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
            Assert.IsTrue(XLColor.Black == XLColor.Black);
        }

        [Test]
        public void ColorNotEqualOperatorInPlace()
        {
            Assert.IsFalse(XLColor.Black != XLColor.Black);
        }

        [Test]
        public void ColorNamedVsHTML()
        {
            Assert.IsTrue(XLColor.Black == XLColor.FromHtml("#000000"));
        }

        [Test]
        public void DefaultColorIndex64isTransparentWhite()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            XLColor color = ws.FirstCell().Style.Fill.BackgroundColor;
            Assert.AreEqual(XLColorType.Indexed, color.ColorType);
            Assert.AreEqual(64, color.Indexed);
            Assert.AreEqual(Color.Transparent, color.Color);
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

            Assert.AreEqual("FFFF0000", color1.Rgb.Value);
            Assert.IsNull(color1.Indexed);
            Assert.IsNull(color1.Theme);
            Assert.IsNull(color1.Tint);

            Assert.IsNull(color2.Rgb);
            Assert.AreEqual(20, color2.Indexed.Value);
            Assert.IsNull(color2.Theme);
            Assert.IsNull(color2.Tint);

            Assert.IsNull(color3.Rgb);
            Assert.IsNull(color3.Indexed);
            Assert.AreEqual(4, color3.Theme.Value);
            Assert.IsNull(color3.Tint);

            Assert.IsNull(color4.Rgb);
            Assert.IsNull(color4.Indexed);
            Assert.AreEqual(5, color4.Theme.Value);
            Assert.AreEqual(0.4, color4.Tint.Value);
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

            Assert.AreEqual("FFFF0000", color1.Rgb.Value);
            Assert.IsNull(color1.Indexed);
            Assert.IsNull(color1.Theme);
            Assert.IsNull(color1.Tint);

            Assert.IsNull(color2.Rgb);
            Assert.AreEqual(20, color2.Indexed.Value);
            Assert.IsNull(color2.Theme);
            Assert.IsNull(color2.Tint);

            Assert.IsNull(color3.Rgb);
            Assert.IsNull(color3.Indexed);
            Assert.AreEqual(4, color3.Theme.Value);
            Assert.IsNull(color3.Tint);

            Assert.IsNull(color4.Rgb);
            Assert.IsNull(color4.Indexed);
            Assert.AreEqual(5, color4.Theme.Value);
            Assert.AreEqual(0.4, color4.Tint.Value);
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

            Assert.AreEqual(XLColorType.Color, xlColor1.ColorType);
            Assert.AreEqual(XLColor.Red.Color, xlColor1.Color);

            Assert.AreEqual(XLColorType.Indexed, xlColor2.ColorType);
            Assert.AreEqual(20, xlColor2.Indexed);

            Assert.AreEqual(XLColorType.Theme, xlColor3.ColorType);
            Assert.AreEqual(XLThemeColor.Accent1, xlColor3.ThemeColor);
            Assert.AreEqual(0, xlColor3.ThemeTint, XLHelper.Epsilon);

            Assert.AreEqual(XLColorType.Theme, xlColor4.ColorType);
            Assert.AreEqual(XLThemeColor.Accent1, xlColor4.ThemeColor);
            Assert.AreEqual(0.4, xlColor4.ThemeTint, XLHelper.Epsilon);
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

            Assert.AreEqual(XLColorType.Color, xlColor1.ColorType);
            Assert.AreEqual(XLColor.Red.Color, xlColor1.Color);

            Assert.AreEqual(XLColorType.Indexed, xlColor2.ColorType);
            Assert.AreEqual(20, xlColor2.Indexed);

            Assert.AreEqual(XLColorType.Theme, xlColor3.ColorType);
            Assert.AreEqual(XLThemeColor.Accent1, xlColor3.ThemeColor);
            Assert.AreEqual(0, xlColor3.ThemeTint, XLHelper.Epsilon);

            Assert.AreEqual(XLColorType.Theme, xlColor4.ColorType);
            Assert.AreEqual(XLThemeColor.Accent1, xlColor4.ThemeColor);
            Assert.AreEqual(0.4, xlColor4.ThemeTint, XLHelper.Epsilon);
        }
    }
}
