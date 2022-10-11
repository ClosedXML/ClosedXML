using ClosedXML.Excel;
using ClosedXML.Graphics;
using NUnit.Framework;
using System;

namespace ClosedXML.SixLabors.Tests
{
    [TestFixture]
    public class FontTests
    {
        private readonly IXLGraphicEngine _engine = SixLaborsEngine.Instance;

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

            public XLFontUnderlineValues Underline
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public XLFontVerticalTextAlignmentValues VerticalAlignment
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public bool Shadow
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public XLColor FontColor
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public XLFontFamilyNumberingValues FontFamilyNumbering
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public XLFontCharSet FontCharSet
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }
        }
    }
}
