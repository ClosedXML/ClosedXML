using System;

namespace ClosedXML.Excel
{
    public interface IXLWithRichString
    {
        IXLRichString AddText(String text);
        IXLRichString AddNewLine();
    }
    public interface IXLRichString: IXLFontBase, IEquatable<IXLRichString>, IXLWithRichString
    {
        String Text { get; set; }


        IXLRichString SetBold(); IXLRichString SetBold(Boolean value);
        IXLRichString SetItalic(); IXLRichString SetItalic(Boolean value);
        IXLRichString SetUnderline(); IXLRichString SetUnderline(XLFontUnderlineValues value);
        IXLRichString SetStrikethrough(); IXLRichString SetStrikethrough(Boolean value);
        IXLRichString SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
        IXLRichString SetShadow(); IXLRichString SetShadow(Boolean value);
        IXLRichString SetFontSize(Double value);
        IXLRichString SetFontColor(XLColor value);
        IXLRichString SetFontName(String value);
        IXLRichString SetFontFamilyNumbering(XLFontFamilyNumberingValues value);
        IXLRichString SetFontCharSet(XLFontCharSet value);
    }
}
