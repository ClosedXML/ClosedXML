using ClosedXML.Excel;
using ClosedXML.Extensions;
using NUnit.Framework;
using SkiaSharp;
using System.Collections.Generic;

namespace ClosedXML.Tests.Extensions
{
    public class FontBaseExtensionsTests
    {
        private const string FontAvaliableOnMostOs = "DejaVu Serif";

        [Test]
        [Platform("Win", Reason = "Expectation only fits windows system because the font calibri isn't available on other OSs")]
        [TestCase(20, 26.25, 3)]
        [TestCase(150, 164, 10)]
        public void ShouldGetHeightUsingCalibriFont(int fontSize, double expectedHeight, int toleratedDiff)
        {
            var fontCache = new Dictionary<IXLFontBase, SKFont>();

            var xLFont = new XLFont
            {
                FontSize = fontSize,
                FontName = "Calibri"
            };

            var actualHeight = xLFont.GetHeight(fontCache);
            Assert.AreEqual(expectedHeight, actualHeight, toleratedDiff);
        }

        [Test]
        [TestCase(20, 27.5, 3)]
        [TestCase(200, 249.75d, 10)]
        public void ShouldGetHeightUsingOsAgnosticFriendlyFont(int fontSize, double expectedHeight, int toleratedDiff)
        {
            var fontCache = new Dictionary<IXLFontBase, SKFont>();

            var xLFont = new XLFont
            {
                FontSize = fontSize,
                FontName = FontAvaliableOnMostOs
            };

            var actualHeight = xLFont.GetHeight(fontCache);
            Assert.AreEqual(expectedHeight, actualHeight, toleratedDiff);
        }

        [Test]
        [Platform("Win", Reason = "Expectation only fits windows system because the font calibri isn't available on other OSs")]
        [TestCase(200, "X", 29.57)]
        [TestCase(20, "Very Wide Column", 30.43)]
        [TestCase(72, "BigText", 45.71)]
        [TestCase(8, "SmallText", 6.86)]
        [TestCase(11, "LongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongText", 241)]
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
        [TestCase(200, "X", 36.29, 2)]
        [TestCase(20, "Very Wide Column", 34.29, 2)]
        [TestCase(72, "BigText", 51.29, 5)]
        [TestCase(8, "SmallText", 8.9, 2)]
        [TestCase(11, "LongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLongTextLong", 228.43, 12)]
        public void ShouldGetWidthUsingOsAgnosticFriendlyFont(int fontSize, string text, double expectedFontSize, int tolerance)
        {
            var fontCache = new Dictionary<IXLFontBase, SKFont>();

            var xLFont = new XLFont
            {
                FontSize = fontSize,
                FontName = FontAvaliableOnMostOs
            };

            var actualWidth = xLFont.GetWidth(text, fontCache);
            Assert.AreEqual(expectedFontSize, actualWidth, tolerance);
        }
    }
}