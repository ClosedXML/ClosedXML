using ClosedXML.Excel;
using ClosedXML.Graphics;
using NUnit.Framework;

namespace ClosedXML.Tests.Graphics
{
    [TestFixture]
    public class FontTests
    {
        private readonly IXLGraphicEngine _engine = DefaultGraphicEngine.Instance.Value;

        [TestCase]
        public void CalculatedTextWidth()
        {
            var textFont = new DummyFont("Calibri", 20);
            var textWidthPt = _engine.GetTextWidth("Lorem ipsum dolor sit amet", textFont, 96);
            Assert.That(textWidthPt, Is.EqualTo(300));
        }

        [TestCase]
        public void CalculatedTextHeight()
        {
            var textFont = new DummyFont("Calibri", 300);
            var textHeightPx = _engine.GetTextHeight(textFont, 96);
            Assert.That(textHeightPx, Is.EqualTo(500));
        }

        [TestCase]
        public void GetMaxDigitWidth()
        {
            var textFont = new DummyFont("Calibri", 11);
            var textWidthPx = _engine.GetMaxDigitWidth(textFont, 96);
            Assert.That(textWidthPx, Is.EqualTo(7.43359375d)); // Calibri,11 has a max digit width of 7 per spec 18.3.1.13
        }

        [TestCase]
        public void DescentIsPositive()
        {
            var textFont = new DummyFont("Calibri", 11);
            var textWidthPt = _engine.GetDescent(textFont, 96);
            Assert.That(textWidthPt, Is.EqualTo(3.666666666666667d));
        }

        [TestCase]
        public void NonExistentFontUsesFallback()
        {
            var nonExistentFont = new DummyFont("NonExistentFont", 100);
            var fallbackFont = new DummyFont("Microsoft Sans Serif", 100);

            var nonExistentFontWidth = _engine.GetTextWidth("ABCDEF text", nonExistentFont, 96);
            var fallbackFontWidth = _engine.GetTextWidth("ABCDEF text", fallbackFont, 96);
            Assert.That(nonExistentFontWidth, Is.EqualTo(fallbackFontWidth));

            var nonExistentFontHeight = _engine.GetTextHeight(nonExistentFont, 96);
            var fallbackFontHeight = _engine.GetTextHeight(fallbackFont, 96);
            Assert.That(nonExistentFontHeight, Is.EqualTo(fallbackFontHeight));
        }

        [TestCase]
        public void CanSpecifyFallbackFontWithoutFileSystem()
        {
            using var fallbackFontStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
            var engine = new DefaultGraphicEngine(fallbackFontStream);

            var nonExistentFont = new DummyFont("Nonexistent Font", 20);
            var widthOfLetterA = engine.GetTextWidth("A", nonExistentFont, 120);

            const double expectedWidthOfLetterA = 31.25d;
            Assert.AreEqual(expectedWidthOfLetterA, widthOfLetterA, 0.0001);
        }

        [TestCase]
        public void CanSpecifyExtraFontsAsStreamsWithoutFileSystem()
        {
            using var fallbackFontStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
            var fontBStream = TestHelper.GetStreamFromResource("Fonts.TestFontB.ttf");
            var engine = new DefaultGraphicEngine(fallbackFontStream, fontBStream);

            var widthOfLetterB = engine.GetTextWidth("B", new DummyFont("TestFontB", 30), 96);

            const double expectedWidthOfLetterB = 25d;
            Assert.AreEqual(expectedWidthOfLetterB, widthOfLetterB, 0.0001);
        }

        private class DummyFont : IXLFontBase
        {
            public DummyFont(string name, double size)
            {
                FontName = name;
                FontSize = size;
            }

            public string FontName { get; set; }

            public double FontSize { get; set; }

            public bool Bold { get; set; }

            public bool Italic { get; set; }

            public bool Strikethrough { get; set; }

            public XLFontUnderlineValues Underline { get; set; } = XLFontUnderlineValues.None;

            public XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }

            public bool Shadow { get; set; }

            public XLColor FontColor { get; set; } = XLColor.Black;

            public XLFontFamilyNumberingValues FontFamilyNumbering { get; set; } = XLFontFamilyNumberingValues.NotApplicable;

            public XLFontCharSet FontCharSet { get; set; } = XLFontCharSet.Default;
        }
    }
}
