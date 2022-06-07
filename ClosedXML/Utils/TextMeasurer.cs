// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using System;
using System.Collections.Generic;

namespace ClosedXML.Utils
{
    internal interface ITextMeasurer : IDisposable
    {
        public System.Drawing.SizeF MeasureString(string text, IXLFontBase fontBase);
    }
}

#if NETFRAMEWORK
namespace ClosedXML.Utils
{
    using System.Drawing;

    internal class GdiPlusMeasurer : ITextMeasurer
    {
        private const int _digits = 2;
        private const decimal _resize = 10000;
        private static readonly StringFormat defaultStringFormat;
        private readonly Dictionary<IXLFontBase, Font> _cache = new Dictionary<IXLFontBase, Font>();

        static GdiPlusMeasurer()
        {
            defaultStringFormat = StringFormat.GenericTypographic;
            defaultStringFormat.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces;
        }

        public SizeF MeasureString(string text, IXLFontBase fontBase)
        {
            // Upscale font and get bounding box of upscaled version of text. Although upscaling doesn't improve
            // float precision fonts are mangled at small sizes.
            var font = GetCachedFont(fontBase);
            var textSize = GraphicsUtils.Graphics.MeasureString(text, font, Int32.MaxValue, defaultStringFormat);
            var width = (float)((decimal)textSize.Width / _resize);
            var height = (float)((decimal)textSize.Height / _resize);
            width = (float)Math.Round(width, _digits, MidpointRounding.AwayFromZero);
            height = (float)Math.Round(height, _digits, MidpointRounding.AwayFromZero);
            return new SizeF(width, height);
        }

        private Font GetCachedFont(IXLFontBase fontBase)
        {
            if (!_cache.TryGetValue(fontBase, out Font font))
            {
                font = new Font(fontBase.FontName, (float)((decimal)fontBase.FontSize * _resize), GetFontStyle(fontBase));
                _cache.Add(fontBase, font);
            }
            return font;
        }

        private static FontStyle GetFontStyle(IXLFontBase font)
        {
            FontStyle fontStyle = FontStyle.Regular;
            if (font.Bold) fontStyle |= FontStyle.Bold;
            if (font.Italic) fontStyle |= FontStyle.Italic;
            if (font.Strikethrough) fontStyle |= FontStyle.Strikeout;
            if (font.Underline != XLFontUnderlineValues.None) fontStyle |= FontStyle.Underline;
            return fontStyle;
        }

        private void DisposeManaged()
        {
            foreach (IDisposable font in _cache.Values)
            {
                font.Dispose();
            }
        }

#if NET40
        public void Dispose()
        {
            // net40 doesn't support Janitor.Fody, so let's dispose manually
            DisposeManaged();
        }

#else
        public void Dispose()
        {
            // Leave this empty (for non net40 targets) so that Janitor.Fody can do its work
        }
#endif
    }
}
#else

namespace ClosedXML.Utils
{
    using SixLabors.Fonts;
    using SizeF = System.Drawing.SizeF;

    internal class SixLaborsFontMeasurer : ITextMeasurer
    {
        // TODO: A better way, like an option or something
        private const string _defaultFallbackFont = "Microsoft Sans Serif";
        private const int _digits = 2;
        private const float _resize = 10000f;
        private readonly Dictionary<IXLFontBase, Font> _cache = new Dictionary<IXLFontBase, Font>();

        public SizeF MeasureString(string text, IXLFontBase fontBase)
        {
            // Upscale font and downsize result, because TextMeasurer.Measure method is using converting
            // decimal bounding box to int before returning.
            var font = GetCachedFont(fontBase);
            var options = new TextOptions(font)
            {
                Dpi = 96,
                KerningMode = KerningMode.Normal
            };

            var textBounds = TextMeasurer.Measure(text, options);

            var width = textBounds.Width / _resize;
            var height = textBounds.Height / _resize;
            width = (float)Math.Round(width, _digits, MidpointRounding.AwayFromZero);
            height = (float)Math.Round(height, _digits, MidpointRounding.AwayFromZero);
            var result = new SizeF(width, height);
            return result;
        }

        private Font GetCachedFont(IXLFontBase fontBase)
        {
            if (!_cache.TryGetValue(fontBase, out var font))
            {
                if (!SystemFonts.TryGet(fontBase.FontName, out var fontFamily))
                {
                    fontFamily = SystemFonts.Get(_defaultFallbackFont);
                }

                font = fontFamily.CreateFont((float)(fontBase.FontSize * _resize), GetFontStyle(fontBase));
                _cache.Add(fontBase, font);
            }
            return font;
        }

        private static FontStyle GetFontStyle(IXLFontBase font)
        {
            var fontStyle = FontStyle.Regular;
            if (font.Bold) fontStyle |= FontStyle.Bold;
            if (font.Italic) fontStyle |= FontStyle.Italic;
            return fontStyle;
        }

        private static TextAttributes GetTextAttributes(IXLFontBase font)
        {
            return font.VerticalAlignment switch
            {
                XLFontVerticalTextAlignmentValues.Baseline => TextAttributes.None,
                XLFontVerticalTextAlignmentValues.Superscript => TextAttributes.Superscript,
                XLFontVerticalTextAlignmentValues.Subscript => TextAttributes.Subscript,
                _ => throw new NotImplementedException()
            };
        }

        private static TextDecorations GetTextDecoration(IXLFontBase font)
        {
            var decorations = TextDecorations.None;
            if (font.Strikethrough) decorations |= TextDecorations.Strikeout;
            if (font.Underline != XLFontUnderlineValues.None) decorations |= TextDecorations.Underline;
            return decorations;
        }

        public void Dispose()
        {
            // TODO: Shoud I dispose of fonts? SixLabors.Fonts doesn't say or show how to dispose, since it's unmanaged, probably not.
        }
    }
}
#endif


