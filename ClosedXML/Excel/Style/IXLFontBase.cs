using System;

namespace ClosedXML.Excel
{
    public interface IXLFontBase
    {
        Boolean Bold { get; set; }

        Boolean Italic { get; set; }

        XLFontUnderlineValues Underline { get; set; }

        Boolean Strikethrough { get; set; }

        XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }

        Boolean Shadow { get; set; }

        Double FontSize { get; set; }

        XLColor FontColor { get; set; }

        String FontName { get; set; }

        XLFontFamilyNumberingValues FontFamilyNumbering { get; set; }

        XLFontCharSet FontCharSet { get; set; }
    }
}
