using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLRichText : IEnumerable<IXLRichString>
    {
        Boolean Bold { set; }
        Boolean Italic { set; }
        XLFontUnderlineValues Underline { set; }
        Boolean Strikethrough { set; }
        XLFontVerticalTextAlignmentValues VerticalAlignment { set; }
        Boolean Shadow { set; }
        Double FontSize { set; }
        IXLColor FontColor { set; }
        String FontName { set; }
        XLFontFamilyNumberingValues FontFamilyNumbering { set; }

        IXLRichText SetBold(); IXLRichText SetBold(Boolean value);
        IXLRichText SetItalic(); IXLRichText SetItalic(Boolean value);
        IXLRichText SetUnderline(); IXLRichText SetUnderline(XLFontUnderlineValues value);
        IXLRichText SetStrikethrough(); IXLRichText SetStrikethrough(Boolean value);
        IXLRichText SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
        IXLRichText SetShadow(); IXLRichText SetShadow(Boolean value);
        IXLRichText SetFontSize(Double value);
        IXLRichText SetFontColor(IXLColor value);
        IXLRichText SetFontName(String value);
        IXLRichText SetFontFamilyNumbering(XLFontFamilyNumberingValues value);

        IXLRichString AddText(String text);
        IXLRichString AddText(String text, IXLFontBase font);
        IXLRichText Clear();
        IXLRichText Substring(Int32 index);
        IXLRichText Substring(Int32 index, Int32 length);
        Int32 Count { get; }
        Int32 Length { get; }


    }
}
