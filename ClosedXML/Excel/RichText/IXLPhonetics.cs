using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLPhoneticAlignment { Center = 0, Distributed = 1, Left = 2, NoControl = 3 }
    public enum XLPhoneticType { FullWidthKatakana = 0, HalfWidthKatakana = 1, Hiragana = 2, NoConversion = 3 }
    public interface IXLPhonetics : IXLFontBase, IEnumerable<IXLPhonetic>, IEquatable<IXLPhonetics>
    {
        IXLPhonetics SetBold(); IXLPhonetics SetBold(Boolean value);
        IXLPhonetics SetItalic(); IXLPhonetics SetItalic(Boolean value);
        IXLPhonetics SetUnderline(); IXLPhonetics SetUnderline(XLFontUnderlineValues value);
        IXLPhonetics SetStrikethrough(); IXLPhonetics SetStrikethrough(Boolean value);
        IXLPhonetics SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
        IXLPhonetics SetShadow(); IXLPhonetics SetShadow(Boolean value);
        IXLPhonetics SetFontSize(Double value);
        IXLPhonetics SetFontColor(XLColor value);
        IXLPhonetics SetFontName(String value);
        IXLPhonetics SetFontFamilyNumbering(XLFontFamilyNumberingValues value);
        IXLPhonetics SetFontCharSet(XLFontCharSet value);

        IXLPhonetics Add(String text, Int32 start, Int32 end);
        IXLPhonetics ClearText();
        IXLPhonetics ClearFont();
        Int32 Count { get; }

        XLPhoneticAlignment Alignment { get; set; }
        XLPhoneticType Type { get; set; }

        IXLPhonetics SetAlignment(XLPhoneticAlignment phoneticAlignment);
        IXLPhonetics SetType(XLPhoneticType phoneticType);
    }
}
