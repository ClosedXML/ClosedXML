using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.Style
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

    public class XLFont
    {
        #region Properties

        private XLRange range;

        private Boolean bold;
        public Boolean Bold
        {
            get
            {
                return bold;
            }
            set
            {
                this.bold = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Font.bold = value);
            }
        }

        private Boolean italic;
        public Boolean Italic
        {
            get
            {
                return italic;
            }
            set
            {
                this.italic = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Font.italic = value);
            }
        }

        private XLFontUnderlineValues underline;
        public XLFontUnderlineValues Underline
        {
            get
            {
                return underline;
            }
            set
            {
                this.underline = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Font.underline = value);
            }
        }

        private Boolean strikethrough;
        public Boolean Strikethrough
        {
            get
            {
                return strikethrough;
            }
            set
            {
                this.strikethrough = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Font.strikethrough = value);
            }
        }

        private XLFontVerticalTextAlignmentValues verticalAlignment;
        public XLFontVerticalTextAlignmentValues VerticalAlignment
        {
            get
            {
                return verticalAlignment;
            }
            set
            {
                this.verticalAlignment = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Font.verticalAlignment = value);
            }
        }

        private Boolean shadow;
        public Boolean Shadow
        {
            get
            {
                return shadow;
            }
            set
            {
                this.shadow = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Font.shadow = value);
            }
        }

        private Double fontSize;
        public Double FontSize
        {
            get
            {
                return fontSize;
            }
            set
            {
                this.fontSize = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Font.fontSize = value);
            }
        }

        private String color;
        public String Color
        {
            get
            {
                return color;
            }
            set
            {
                this.color = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Font.color = value);

            }
        }

        private String fontName;
        public String FontName
        {
            get
            {
                return fontName;
            }
            set
            {
                this.fontName = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Font.fontName = value);
            }
        }

        private Int32 fontFamilyNumbering;
        public Int32 FontFamilyNumbering
        {
            get
            {
                return fontFamilyNumbering;
            }
            set
            {
                this.fontFamilyNumbering = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Font.fontFamilyNumbering = value);
            }
        }

        #endregion

        #region Constructors

        public XLFont(XLFont defaultFont, XLRange range)
        {
            this.range = range;
            if (defaultFont != null)
            {
                Bold = defaultFont.Bold;
                Italic = defaultFont.Italic;
                Underline = defaultFont.Underline;
                Strikethrough = defaultFont.Strikethrough;
                VerticalAlignment = defaultFont.VerticalAlignment;
                Shadow = defaultFont.Shadow;
                FontFamilyNumbering = defaultFont.FontFamilyNumbering;
                FontName = defaultFont.FontName;
                FontSize = defaultFont.FontSize;
                Color = defaultFont.Color;
            }
        }

        #endregion

        #region Overridden

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(Bold.ToString());
            sb.Append("-");
            sb.Append(FontSize.ToString());
            sb.Append("-");
            sb.Append(Color);
            sb.Append("-");
            sb.Append(FontName);
            sb.Append("-");
            sb.Append(FontFamilyNumbering.ToString());
            return sb.ToString();
        }

        #endregion
    }
}
