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
        IXLPhonetics SetFontScheme(XLFontScheme value);

        /// <summary>
        /// Add a phonetic run above a base text. Phonetic runs can't overlap.
        /// </summary>
        /// <param name="text">Text to display above a section of a base text. Can't be empty.</param>
        /// <param name="start">Index of a first character of a base  text above which should <paramref name="text"/> be displayed. Valid values are <c>0</c>..<c>length-1</c>.</param>
        /// <param name="end">The excluded ending index in a base text (the hint is not displayed above the <c>end</c>). Must be &gt; <paramref name="start"/>. Valid values are <c>1</c>..<c>length</c>.</param>
        IXLPhonetics Add(String text, Int32 start, Int32 end);

        /// <summary>
        /// Remove all phonetic runs. Keeps font properties.
        /// </summary>
        IXLPhonetics ClearText();

        /// <summary>
        /// Reset font properties to the default font of a container (likely <c>IXLCell</c>). Keeps phonetic runs, <see cref="Type"/> and <see cref="Alignment"/>.
        /// </summary>
        IXLPhonetics ClearFont();

        /// <summary>
        /// Number of phonetic runs above the base text.
        /// </summary>
        Int32 Count { get; }

        XLPhoneticAlignment Alignment { get; set; }
        XLPhoneticType Type { get; set; }

        IXLPhonetics SetAlignment(XLPhoneticAlignment phoneticAlignment);
        IXLPhonetics SetType(XLPhoneticType phoneticType);
    }
}
