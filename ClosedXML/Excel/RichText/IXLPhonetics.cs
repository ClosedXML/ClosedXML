using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLPhoneticAlignment { Center, Distributed, Left, NoControl }
    public enum XLPhoneticType { FullWidthKatakana, HalfWidthKatakana, Hiragana, NoConversion }
    public interface IXLPhonetics : IXLFontBase, IEnumerable<IXLPhonetic>, IEquatable<IXLPhonetics>
    {
        IXLPhonetics SetBold(); IXLPhonetics SetBold(bool value);
        IXLPhonetics SetItalic(); IXLPhonetics SetItalic(bool value);
        IXLPhonetics SetUnderline(); IXLPhonetics SetUnderline(XLFontUnderlineValues value);
        IXLPhonetics SetStrikethrough(); IXLPhonetics SetStrikethrough(bool value);
        IXLPhonetics SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
        IXLPhonetics SetShadow(); IXLPhonetics SetShadow(bool value);
        IXLPhonetics SetFontSize(double value);
        IXLPhonetics SetFontColor(XLColor value);
        IXLPhonetics SetFontName(string value);
        IXLPhonetics SetFontFamilyNumbering(XLFontFamilyNumberingValues value);
        IXLPhonetics SetFontCharSet(XLFontCharSet value);

        IXLPhonetics Add(string text, int start, int end);
        IXLPhonetics ClearText();
        IXLPhonetics ClearFont();
        int Count { get; }

        XLPhoneticAlignment Alignment { get; set; }
        XLPhoneticType Type { get; set; }

        IXLPhonetics SetAlignment(XLPhoneticAlignment phoneticAlignment);
        IXLPhonetics SetType(XLPhoneticType phoneticType);
    }
}
