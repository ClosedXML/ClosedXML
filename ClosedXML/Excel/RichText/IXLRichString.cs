using System;

namespace ClosedXML.Excel
{
    public interface IXLWithRichString
    {
        IXLRichString AddText(String text);
        IXLRichString AddNewLine();
    }
    public interface IXLRichString : IXLFontBase, IEquatable<IXLRichString>, IXLWithRichString
    {
        String Text { get; set; }

        /// <summary>
        /// Determines a theme scheme the rich strings belongs to. If the string belongs
        /// to a scheme and user changes theme in Excel, the font of the string will switch
        /// to the new theme font.
        /// </summary>
        XLFontScheme FontScheme { get; set; }

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

        /// <inheritdoc cref="FontScheme"/>
        IXLRichString SetFontScheme(XLFontScheme value);
    }
}
