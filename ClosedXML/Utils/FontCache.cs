// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace ClosedXML.Utils
{
    internal class FontCache : IDisposable
    {
        private static readonly StringFormat defaultStringFormat = StringFormat.GenericTypographic;
        private readonly Dictionary<IXLFontBase, Font> _cache = new Dictionary<IXLFontBase, Font>();

        public SizeF MeasureString(string text, IXLFontBase fontBase)
        {
            var font = GetCachedFont(fontBase);
            SizeF result = GraphicsUtils.Graphics.MeasureString(text, font, Int32.MaxValue, defaultStringFormat);
            return result;
        }

        private Font GetCachedFont(IXLFontBase fontBase)
        {
            if (!_cache.TryGetValue(fontBase, out Font font))
            {
                font = new Font(fontBase.FontName, (float)fontBase.FontSize, GetFontStyle(fontBase));
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
