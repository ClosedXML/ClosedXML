using System;

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

        public Boolean Bold { get; set; }
        public Boolean Italic { get; set; }
        public XLFontUnderlineValues Underline { get; set; }
        public Boolean Strikethrough { get; set; }
        public XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }
        public Boolean Shadow { get; set; }
        public Double FontSize { get; set; }
        public XLColor FontColor { get; set; }
        public String FontName { get; set; }
        public XLFontFamilyNumberingValues FontFamilyNumbering { get; set; }

        public XLFontCharSet FontCharSet { get; set; }

        public IXLDrawingStyle SetBold()
        {
            Bold = true;
            return _style;
        }

        public IXLDrawingStyle SetBold(Boolean value)
        {
            Bold = value;
            return _style;
        }

        public IXLDrawingStyle SetItalic()
        {
            Italic = true;
            return _style;
        }

        public IXLDrawingStyle SetItalic(Boolean value)
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

        public IXLDrawingStyle SetStrikethrough(Boolean value)
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

        public IXLDrawingStyle SetShadow(Boolean value)
        {
            Shadow = value;
            return _style;
        }

        public IXLDrawingStyle SetFontSize(Double value)
        {
            FontSize = value;
            return _style;
        }

        public IXLDrawingStyle SetFontColor(XLColor value)
        {
            FontColor = value;
            return _style;
        }

        public IXLDrawingStyle SetFontName(String value)
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
