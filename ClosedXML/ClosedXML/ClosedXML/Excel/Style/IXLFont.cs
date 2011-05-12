using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ClosedXML.Excel
{
    public enum XLFontUnderlineValues
    {
        Double,
        DoubleAccounting,
        None,
        Single,
        SingleAccounting
    }

    public enum XLFontVerticalTextAlignmentValues
    {
        Baseline,
        Subscript,
        Superscript
    }

    public enum XLFontFamilyNumberingValues
    { 
        NotApplicable = 0,
        Roman = 1,
        Swiss = 2,
        Modern = 3,
        Script = 4,
        Decorative = 5
    }

    public interface IXLFont : IEquatable<IXLFont>
    {
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

        IXLStyle SetBold(); IXLStyle SetBold(Boolean value);
        IXLStyle SetItalic(); IXLStyle SetItalic(Boolean value);
        IXLStyle SetUnderline(); IXLStyle SetUnderline(XLFontUnderlineValues value);
        IXLStyle SetStrikethrough(); IXLStyle SetStrikethrough(Boolean value);
        IXLStyle SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
        IXLStyle SetShadow(); IXLStyle SetShadow(Boolean value);
        IXLStyle SetFontSize(Double value);
        IXLStyle SetFontColor(IXLColor value);
        IXLStyle SetFontName(String value);
        IXLStyle SetFontFamilyNumbering(XLFontFamilyNumberingValues value);

    }
}
