using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLComment: IXLDrawing<IXLComment>, IEnumerable<IXLRichString>, IEquatable<IXLComment>
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

        IXLComment SetBold(); IXLComment SetBold(Boolean value);
        IXLComment SetItalic(); IXLComment SetItalic(Boolean value);
        IXLComment SetUnderline(); IXLComment SetUnderline(XLFontUnderlineValues value);
        IXLComment SetStrikethrough(); IXLComment SetStrikethrough(Boolean value);
        IXLComment SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
        IXLComment SetShadow(); IXLComment SetShadow(Boolean value);
        IXLComment SetFontSize(Double value);
        IXLComment SetFontColor(IXLColor value);
        IXLComment SetFontName(String value);
        IXLComment SetFontFamilyNumbering(XLFontFamilyNumberingValues value);

        IXLRichString AddText(String text);
        IXLRichString AddText(String text, IXLFontBase font);
        IXLComment ClearText();
        IXLComment ClearFont();
        IXLComment Substring(Int32 index);
        IXLComment Substring(Int32 index, Int32 length);
        Int32 Count { get; }
        Int32 Length { get; }

        String Text { get; }
        IXLPhonetics Phonetics { get; }
        Boolean HasPhonetics { get; }
    }
}
