using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLRichString : IEnumerable<IXLRichText>
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

        IXLRichString SetBold(); IXLRichString SetBold(Boolean value);
        IXLRichString SetItalic(); IXLRichString SetItalic(Boolean value);
        IXLRichString SetUnderline(); IXLRichString SetUnderline(XLFontUnderlineValues value);
        IXLRichString SetStrikethrough(); IXLRichString SetStrikethrough(Boolean value);
        IXLRichString SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
        IXLRichString SetShadow(); IXLRichString SetShadow(Boolean value);
        IXLRichString SetFontSize(Double value);
        IXLRichString SetFontColor(IXLColor value);
        IXLRichString SetFontName(String value);
        IXLRichString SetFontFamilyNumbering(XLFontFamilyNumberingValues value);

        IXLRichText AddText(String text);
        IXLRichText AddText(String text, IXLFontBase font);
        IXLRichString Clear();
        IXLRichString Substring(Int32 index);
        IXLRichString Substring(Int32 index, Int32 length);
        Int32 Count { get; }
        Int32 Length { get; }


    }
}
