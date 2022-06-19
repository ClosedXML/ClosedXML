namespace ClosedXML.Excel
{
    internal class XLDrawingFont : IXLDrawingFont
    {
        private readonly IXLDrawingStyle _style;

        public XLDrawingFont(IXLDrawingStyle style)
        {
            _style = style;
            FontName = "Tahoma";
            FontSize = 9;
            Underline = XLFontUnderlineValues.None;
            FontColor = XLColor.FromIndex(64);
        }

        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public XLFontUnderlineValues Underline { get; set; }
        public bool Strikethrough { get; set; }
        public XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }
        public bool Shadow { get; set; }
        public double FontSize { get; set; }
        public XLColor FontColor { get; set; }
        public string FontName { get; set; }
        public XLFontFamilyNumberingValues FontFamilyNumbering { get; set; }

        public XLFontCharSet FontCharSet { get; set; }

        public IXLDrawingStyle SetBold()
        {
            Bold = true;
            return _style;
        }

        public IXLDrawingStyle SetBold(bool value)
        {
            Bold = value;
            return _style;
        }

        public IXLDrawingStyle SetItalic()
        {
            Italic = true;
            return _style;
        }

        public IXLDrawingStyle SetItalic(bool value)
        {
            Italic = value;
            return _style;
        }

        public IXLDrawingStyle SetUnderline()
        {
            Underline = XLFontUnderlineValues.Single;
            return _style;
        }

        public IXLDrawingStyle SetUnderline(XLFontUnderlineValues value)
        {
            Underline = value;
            return _style;
        }

        public IXLDrawingStyle SetStrikethrough()
        {
            Strikethrough = true;
            return _style;
        }

        public IXLDrawingStyle SetStrikethrough(bool value)
        {
            Strikethrough = value;
            return _style;
        }

        public IXLDrawingStyle SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
        {
            VerticalAlignment = value;
            return _style;
        }

        public IXLDrawingStyle SetShadow()
        {
            Shadow = true;
            return _style;
        }

        public IXLDrawingStyle SetShadow(bool value)
        {
            Shadow = value;
            return _style;
        }

        public IXLDrawingStyle SetFontSize(double value)
        {
            FontSize = value;
            return _style;
        }

        public IXLDrawingStyle SetFontColor(XLColor value)
        {
            FontColor = value;
            return _style;
        }

        public IXLDrawingStyle SetFontName(string value)
        {
            FontName = value;
            return _style;
        }

        public IXLDrawingStyle SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
        {
            FontFamilyNumbering = value;
            return _style;
        }

        public IXLDrawingStyle SetFontCharSet(XLFontCharSet value)
        {
            FontCharSet = value;
            return _style;
        }
    }
}
