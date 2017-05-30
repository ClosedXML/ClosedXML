using System;

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

    public interface IXLFont : IXLFontBase, IEquatable<IXLFont>
    {
        IXLStyle SetBold(); IXLStyle SetBold(Boolean value);
        IXLStyle SetItalic(); IXLStyle SetItalic(Boolean value);
        IXLStyle SetUnderline(); IXLStyle SetUnderline(XLFontUnderlineValues value);
        IXLStyle SetStrikethrough(); IXLStyle SetStrikethrough(Boolean value);
        IXLStyle SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
        IXLStyle SetShadow(); IXLStyle SetShadow(Boolean value);
        IXLStyle SetFontSize(Double value);
        IXLStyle SetFontColor(XLColor value);
        IXLStyle SetFontName(String value);
        IXLStyle SetFontFamilyNumbering(XLFontFamilyNumberingValues value);

    }
}
