namespace ClosedXML.Excel
{
    public interface IXLFontBase
    {
        bool Bold { get; set; }

        bool Italic { get; set; }

        XLFontUnderlineValues Underline { get; set; }

        bool Strikethrough { get; set; }

        XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }

        bool Shadow { get; set; }

        double FontSize { get; set; }

        XLColor FontColor { get; set; }

        string FontName { get; set; }

        XLFontFamilyNumberingValues FontFamilyNumbering { get; set; }

        XLFontCharSet FontCharSet { get; set; }
    }
}
