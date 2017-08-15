using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLFormattedText<T> : IEnumerable<IXLRichString>, IEquatable<IXLFormattedText<T>>, IXLWithRichString
    {
        Boolean Bold { set; }
        Boolean Italic { set; }
        XLFontUnderlineValues Underline { set; }
        Boolean Strikethrough { set; }
        XLFontVerticalTextAlignmentValues VerticalAlignment { set; }
        Boolean Shadow { set; }
        Double FontSize { set; }
        XLColor FontColor { set; }
        String FontName { set; }
        XLFontFamilyNumberingValues FontFamilyNumbering { set; }

        IXLFormattedText<T> SetBold(); IXLFormattedText<T> SetBold(Boolean value);
        IXLFormattedText<T> SetItalic(); IXLFormattedText<T> SetItalic(Boolean value);
        IXLFormattedText<T> SetUnderline(); IXLFormattedText<T> SetUnderline(XLFontUnderlineValues value);
        IXLFormattedText<T> SetStrikethrough(); IXLFormattedText<T> SetStrikethrough(Boolean value);
        IXLFormattedText<T> SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
        IXLFormattedText<T> SetShadow(); IXLFormattedText<T> SetShadow(Boolean value);
        IXLFormattedText<T> SetFontSize(Double value);
        IXLFormattedText<T> SetFontColor(XLColor value);
        IXLFormattedText<T> SetFontName(String value);
        IXLFormattedText<T> SetFontFamilyNumbering(XLFontFamilyNumberingValues value);

        IXLRichString AddText(String text, IXLFontBase font);
        IXLFormattedText<T> ClearText();
        IXLFormattedText<T> ClearFont();
        IXLFormattedText<T> Substring(Int32 index);
        IXLFormattedText<T> Substring(Int32 index, Int32 length);
        Int32 Count { get; }
        Int32 Length { get; }

        String Text { get; }
        IXLPhonetics Phonetics { get; }
        Boolean HasPhonetics { get; }
    }
}
