using ClosedXML.Excel;
using ClosedXML.Extensions;
using NUnit.Framework;
using SkiaSharp;
using System.Collections.Generic;

namespace ClosedXML.Tests.Extensions
{
    public class FontBaseExtensionsTests
    {
        [Test]
        [Platform("Win", Reason = "Expectation only fits windows system because the font calibri isn't available on other OSs")]
        public void ShouldGetHeightUsingCalibriFont()
        {
            var fontCache = new Dictionary<IXLFontBase, SKFont>();

            var xLFont = new XLFont
            {
                FontSize = 200,
                FontName = "Calibri"
            };

            var actualHeight = xLFont.GetHeight(fontCache);
            Assert.AreEqual(255, actualHeight, 10);
        }

        [Test]
        public void ShouldGetHeightUsingOsAgnosticFriendlyFont()
        {
            var fontCache = new Dictionary<IXLFontBase, SKFont>();

            var xLFont = new XLFont
            {
                FontSize = 200,
                FontName = "Verdana"
            };

            var actualHeight = xLFont.GetHeight(fontCache);
            Assert.AreEqual(288, actualHeight, 10);
        }

        [Test]
        [Platform("Win", Reason = "Expectation only fits windows system because the font calibri isn't available on other OSs")]
        [TestCase(200, "X", 28)]
        [TestCase(20, "Very Wide Column", 28.55)]
        [TestCase(72, "BigText", 43.64)]
        [TestCase(8, "SmallText", 6.55)]
        // Excel adjusts to unreasonable large 253.09
        [TestCase(11, "LongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongText", 244.45)]
        public void ShouldGetWidthCalibriFont(int fontSize, string text, double expectedFontSize)
        {
            var fontCache = new Dictionary<IXLFontBase, SKFont>();

            var xLFont = new XLFont
            {
                FontSize = fontSize,
                FontName = "Calibri"
            };

            var actualWidth = xLFont.GetWidth(text, fontCache);
            Assert.AreEqual(expectedFontSize, actualWidth, 3);
        }

        [Test]
        [TestCase(8, "SmallText", 8.05, 2)]
        [TestCase(11, "LongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongText", 245, 12)]
        [TestCase(20, "Very Wide Column", 36.18, 2)]
        [TestCase(72, "BigText", 55.27, 5)]
        [TestCase(200, "X", 37, 5)]
        public void ShouldGetWidthUsingOsAgnosticFriendlyFont(int fontSize, string text, double expectedFontSize, int tolerance)
        {
            var fontCache = new Dictionary<IXLFontBase, SKFont>();

            var xLFont = new XLFont
            {
                FontSize = fontSize,
                FontName = "Verdana"
            };

            var actualWidth = xLFont.GetWidth(text, fontCache);
            Assert.AreEqual(expectedFontSize, actualWidth, tolerance);
        }
    }
}