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

    public enum XLFontCharSet
    {
        /// <summary>
        /// ASCII character set.
        /// </summary>
        Ansi = 0,

        /// <summary>
        /// System default character set.
        /// </summary>
        Default = 1,

        /// <summary>
        /// Symbol character set.
        /// </summary>
        Symbol = 2,

        /// <summary>
        /// Characters used by Macintosh.
        /// </summary>
        Mac = 77,

        /// <summary>
        /// Japanese character set.
        /// </summary>
        ShiftJIS = 128,

        /// <summary>
        /// Korean character set.
        /// </summary>
        Hangul = 129,

        /// <summary>
        /// Another common spelling of the Korean character set.
        /// </summary>
        Hangeul = 129,

        /// <summary>
        /// Korean character set.
        /// </summary>
        Johab = 130,

        /// <summary>
        /// Chinese character set used in mainland China.
        /// </summary>
        GB2312 = 134,

        /// <summary>
        /// Chinese character set used mostly in Hong Kong SAR and Taiwan.
        /// </summary>
        ChineseBig5 = 136,

        /// <summary>
        /// Greek character set.
        /// </summary>
        Greek = 161,

        /// <summary>
        /// Turkish character set.
        /// </summary>
        Turkish = 162,

        /// <summary>
        /// Vietnamese character set.
        /// </summary>
        Vietnamese = 163,

        /// <summary>
        /// Hebrew character set.
        /// </summary>
        Hebrew = 177,

        /// <summary>
        /// Arabic character set.
        /// </summary>
        Arabic = 178,

        /// <summary>
        /// Baltic character set.
        /// </summary>
        Baltic = 186,

        /// <summary>
        /// Russian character set.
        /// </summary>
        Russian = 204,

        /// <summary>
        /// Thai character set.
        /// </summary>
        Thai = 222,

        /// <summary>
        /// Eastern European character set.
        /// </summary>
        EastEurope = 238,

        /// <summary>
        /// Extended ASCII character set used with disk operating system (DOS) and some Microsoft Windows fonts.
        /// </summary>
        Oem = 255
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

        IXLStyle SetFontCharSet(XLFontCharSet value);
    }
}
