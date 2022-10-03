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
            var textWidthPt = _engine.GetTextWidth("Lorem ipsum dolor sit amet", textFont);
            Assert.That(textWidthPt, Is.EqualTo(225));
        }

        [TestCase]
        public void CalculatedTextHeight()
        {
            var textFont = new DummyFont("Calibri", 100);
            var textHeight = _engine.GetTextHeight(textFont);
            Assert.That(textHeight, Is.EqualTo(125));
        }

        [TestCase]
        public void GetMaxDigitWidth()
        {
            var textFont = new DummyFont("Calibri", 11);
            var textWidthPt = _engine.GetMaxDigitWidth(textFont);
            Assert.That(textWidthPt, Is.EqualTo(5.5751953125d));
            Assert.That(Math.Ceiling(textWidthPt / 72d * 96d), Is.EqualTo(8)); // Calibri,11 has a max digit width of 8 per spec
        }

        [TestCase]
        public void AscentPlusDescentIsFontSize()
        {
            var fontSize = 20;
            var textFont = new DummyFont("Calibri", fontSize);
            var emSquareSize = _engine.GetAscent(textFont) + _engine.GetDescent(textFont);
            Assert.That(fontSize, Is.EqualTo(emSquareSize));
        }

        [TestCase]
        public void DescentIsPositive()
        {
            var textFont = new DummyFont("Calibri", 11);
            var textWidthPt = _engine.GetDescent(textFont);
            Assert.That(textWidthPt, Is.EqualTo(2.75));
        }

        [TestCase]
        public void NonExistentFontUsesFallback()
        {
            var nonExistentFont = new DummyFont("NonExistentFont", 100);
            var fallbackFont = new DummyFont("Microsoft Sans Serif", 100);

            var nonExistentFontWidth = _engine.GetTextWidth("ABCDEF text", nonExistentFont);
            var fallbackFontWidth = _engine.GetTextWidth("ABCDEF text", fallbackFont);
            Assert.That(nonExistentFontWidth, Is.EqualTo(fallbackFontWidth));

            var nonExistentFontHeight = _engine.GetTextHeight(nonExistentFont);
            var fallbackFontHeight = _engine.GetTextHeight(fallbackFont);
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
