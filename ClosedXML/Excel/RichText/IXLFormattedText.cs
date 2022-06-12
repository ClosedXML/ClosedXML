using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLFormattedText<T> : IEnumerable<IXLRichString>, IEquatable<IXLFormattedText<T>>, IXLWithRichString
    {
        bool Bold { set; }
        bool Italic { set; }
        XLFontUnderlineValues Underline { set; }
        bool Strikethrough { set; }
        XLFontVerticalTextAlignmentValues VerticalAlignment { set; }
        bool Shadow { set; }
        double FontSize { set; }
        XLColor FontColor { set; }
        string FontName { set; }
        XLFontFamilyNumberingValues FontFamilyNumbering { set; }

        IXLFormattedText<T> SetBold(); IXLFormattedText<T> SetBold(bool value);
        IXLFormattedText<T> SetItalic(); IXLFormattedText<T> SetItalic(bool value);
        IXLFormattedText<T> SetUnderline(); IXLFormattedText<T> SetUnderline(XLFontUnderlineValues value);
        IXLFormattedText<T> SetStrikethrough(); IXLFormattedText<T> SetStrikethrough(bool value);
        IXLFormattedText<T> SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
        IXLFormattedText<T> SetShadow(); IXLFormattedText<T> SetShadow(bool value);
        IXLFormattedText<T> SetFontSize(double value);
        IXLFormattedText<T> SetFontColor(XLColor value);
        IXLFormattedText<T> SetFontName(string value);
        IXLFormattedText<T> SetFontFamilyNumbering(XLFontFamilyNumberingValues value);

        IXLRichString AddText(string text, IXLFontBase font);
        IXLFormattedText<T> ClearText();
        IXLFormattedText<T> ClearFont();
        IXLFormattedText<T> Substring(int index);
        IXLFormattedText<T> Substring(int index, int length);
        int Count { get; }
        int Length { get; }

        string Text { get; }
        IXLPhonetics Phonetics { get; }
        bool HasPhonetics { get; }
    }
}
