using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLRichText: IEquatable<IXLRichText>
    {
        String Text { get; }

        Boolean Bold { get; set; }
        Boolean Italic { get; set; }
        XLFontUnderlineValues Underline { get; set; }
        Boolean Strikethrough { get; set; }
        XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }
        Boolean Shadow { get; set; }
        Double FontSize { get; set; }
        IXLColor FontColor { get; set; }
        String FontName { get; set; }
        XLFontFamilyNumberingValues FontFamilyNumbering { get; set; }

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
    }
}
