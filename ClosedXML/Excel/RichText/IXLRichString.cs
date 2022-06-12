using System;

namespace ClosedXML.Excel
{
    public interface IXLWithRichString
    {
        IXLRichString AddText(string text);
        IXLRichString AddNewLine();
    }
    public interface IXLRichString: IXLFontBase, IEquatable<IXLRichString>, IXLWithRichString
    {
        string Text { get; set; }


        IXLRichString SetBold(); IXLRichString SetBold(bool value);
        IXLRichString SetItalic(); IXLRichString SetItalic(bool value);
        IXLRichString SetUnderline(); IXLRichString SetUnderline(XLFontUnderlineValues value);
        IXLRichString SetStrikethrough(); IXLRichString SetStrikethrough(bool value);
        IXLRichString SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
        IXLRichString SetShadow(); IXLRichString SetShadow(bool value);
        IXLRichString SetFontSize(double value);
        IXLRichString SetFontColor(XLColor value);
        IXLRichString SetFontName(string value);
        IXLRichString SetFontFamilyNumbering(XLFontFamilyNumberingValues value);
        IXLRichString SetFontCharSet(XLFontCharSet value);
    }
}
