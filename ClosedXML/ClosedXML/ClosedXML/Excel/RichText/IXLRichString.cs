using System;

namespace ClosedXML.Excel
{
    public interface IXLRichString: IXLFontBase, IEquatable<IXLRichString>
    {
        String Text { get; }

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
    }
}
