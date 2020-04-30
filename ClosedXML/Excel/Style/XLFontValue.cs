using ClosedXML.Excel.Caching;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal sealed class XLFontValue
    {
        private static readonly XLFontRepository Repository = new XLFontRepository(key => new XLFontValue(key));

        public static XLFontValue FromKey(ref XLFontKey key)
        {
            return Repository.GetOrCreate(ref key);
        }

        private static readonly XLFontKey DefaultKey = new XLFontKey
        {
            Bold = false,
            Italic = false,
            Underline = XLFontUnderlineValues.None,
            Strikethrough = false,
            VerticalAlignment = XLFontVerticalTextAlignmentValues.Baseline,
            FontSize = 11,
            FontColor = XLColor.FromArgb(0, 0, 0).Key,
            FontName = "Calibri",
            FontFamilyNumbering = XLFontFamilyNumberingValues.Swiss,
            FontCharSet = XLFontCharSet.Default
        };
        internal static readonly XLFontValue Default = FromKey(ref DefaultKey);

        public XLFontKey Key { get; private set; }

        public bool Bold { get { return Key.Bold; } }

        public bool Italic { get { return Key.Italic; } }

        public XLFontUnderlineValues Underline { get { return Key.Underline; } }

        public bool Strikethrough { get { return Key.Strikethrough; } }

        public XLFontVerticalTextAlignmentValues VerticalAlignment { get { return Key.VerticalAlignment; } }

        public bool Shadow { get { return Key.Shadow; } }

        public double FontSize { get { return Key.FontSize; } }

        public XLColor FontColor { get; private set; }

        public string FontName { get { return Key.FontName; } }

        public XLFontFamilyNumberingValues FontFamilyNumbering { get { return Key.FontFamilyNumbering; } }

        public XLFontCharSet FontCharSet { get { return Key.FontCharSet; } }

        private XLFontValue(XLFontKey key)
        {
            Key = key;
            var fontColorKey = Key.FontColor;
            FontColor = XLColor.FromKey(ref fontColorKey);
        }

        public override bool Equals(object obj)
        {
            var cached = obj as XLFontValue;
            return cached != null &&
                   Key.Equals(cached.Key);
        }

        public override int GetHashCode()
        {
            return -280332839 + EqualityComparer<XLFontKey>.Default.GetHashCode(Key);
        }
    }
}
