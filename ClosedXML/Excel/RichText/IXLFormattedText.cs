#nullable disable

using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLFormattedText<T> : IEnumerable<IXLRichString>, IEquatable<IXLFormattedText<T>>, IXLWithRichString
    {
        Boolean Bold { set; }
        Boolean Italic { set; }
        XLFontUnderlineValues Underline { set; }
        Boolean Strikethrough { set; }
        XLFontVerticalTextAlignmentValues VerticalAlignment { set; }
        Boolean Shadow { set; }
        Double FontSize { set; }
        XLColor FontColor { set; }
        String FontName { set; }
        XLFontFamilyNumberingValues FontFamilyNumbering { set; }

        IXLFormattedText<T> SetBold(); IXLFormattedText<T> SetBold(Boolean value);
        IXLFormattedText<T> SetItalic(); IXLFormattedText<T> SetItalic(Boolean value);
        IXLFormattedText<T> SetUnderline(); IXLFormattedText<T> SetUnderline(XLFontUnderlineValues value);
        IXLFormattedText<T> SetStrikethrough(); IXLFormattedText<T> SetStrikethrough(Boolean value);
        IXLFormattedText<T> SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
        IXLFormattedText<T> SetShadow(); IXLFormattedText<T> SetShadow(Boolean value);
        IXLFormattedText<T> SetFontSize(Double value);
        IXLFormattedText<T> SetFontColor(XLColor value);
        IXLFormattedText<T> SetFontName(String value);
        IXLFormattedText<T> SetFontFamilyNumbering(XLFontFamilyNumberingValues value);

        IXLRichString AddText(String text, IXLFontBase font);
        IXLFormattedText<T> ClearText();
        IXLFormattedText<T> ClearFont();
        IXLFormattedText<T> Substring(Int32 index);
        IXLFormattedText<T> Substring(Int32 index, Int32 length);

        /// <summary>
        /// Replace the text and formatting of this text by texts and formatting from the <paramref name="original"/> text.
        /// </summary>
        /// <param name="original">Original to copy from.</param>
        /// <returns>This text.</returns>
        IXLFormattedText<T> CopyFrom(IXLFormattedText<T> original);

        /// <summary>
        /// How many rich strings is the formatted text composed of.
        /// </summary>
        Int32 Count { get; }

        /// <summary>
        /// Length of the whole formatted text.
        /// </summary>
        Int32 Length { get; }

        /// <summary>
        /// Get text of the whole formatted text.
        /// </summary>
        String Text { get; }

        /// <summary>
        /// Does this text has phonetics? Unlike accessing the <see cref="Phonetics"/> property, this method
        /// doesn't create a new instance on access.
        /// </summary>
        Boolean HasPhonetics { get; }

        /// <summary>
        /// Get or create phonetics for the text. Use <see cref="HasPhonetics"/> to check for existence to avoid unnecessary creation.
        /// </summary>
        IXLPhonetics Phonetics { get; }
    }
}
