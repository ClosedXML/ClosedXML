using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ClosedXML.Excel
{
    public partial class XLColor: IXLColor
    {
        public Boolean HasValue { get; private set; }

        XLColorType colorType;
        public XLColorType ColorType {
            get
            {
                return colorType;
            }
            private set
            {
                colorType = value;
            }
        }
        private Color color;
        public Color Color 
        {
            get
            {
                if (colorType == XLColorType.Theme)
                {
                    //if (workbook == null)
                        throw new Exception("Cannot convert theme color to Color.");
                    //else
                    //    return workbook.GetXLColor(themeColor).Color;
                }
                else if (colorType == XLColorType.Indexed)
                {
                    return IndexedColors[indexed].Color;
                }
                else
                {
                    return color;
                }
            }
            private set
            {
                color = value;
                colorType = XLColorType.Color;
            }
        }

        private Int32 indexed;
        public Int32 Indexed
        {
            get
            {
                if (ColorType == XLColorType.Theme)
                {
                    throw new Exception("Cannot convert theme color to indexed color.");
                }
                else if (ColorType == XLColorType.Indexed)
                {
                    return indexed;
                }
                else // ColorType == Color
                {
                    throw new Exception("Cannot convert Color to indexed color.");
                }
            }
            private set
            {
                indexed = value;
                colorType = XLColorType.Indexed;
            }
        }

        private XLThemeColor themeColor;
        public XLThemeColor ThemeColor
        {
            get
            {
                if (ColorType == XLColorType.Theme)
                {
                    return themeColor;
                }
                else if (ColorType == XLColorType.Indexed)
                {
                    throw new Exception("Cannot convert indexed color to theme color.");
                }
                else // ColorType == Color
                {
                    throw new Exception("Cannot convert Color to theme color.");
                }
            }
            private set
            {
                themeColor = value;
                if (themeTint == 0)
                    themeTint = 1;
                colorType = XLColorType.Theme;
            }
        }

        private Double themeTint;
        public Double ThemeTint
        {
            get
            {
                if (ColorType == XLColorType.Theme)
                {
                    return themeTint;
                }
                else if (ColorType == XLColorType.Indexed)
                {
                    throw new Exception("Cannot extract theme tint from an indexed color.");
                }
                else // ColorType == Color
                {
                    return (Double)color.A / 255.0;
                }
            }
            private set
            {
                themeTint = value;
                colorType = XLColorType.Theme;
            }
        }

        public bool Equals(IXLColor other)
        {
            var otherC = other as XLColor;
            if (colorType == otherC.colorType)
            {
                if (colorType == XLColorType.Color)
                {
                    return color.ToArgb() == otherC.color.ToArgb();
                }
                if (colorType == XLColorType.Theme)
                {
                    return themeColor == otherC.themeColor
                        && themeTint == otherC.themeTint;
                }
                else
                {
                    return indexed == otherC.indexed;
                }
            }
            else
            {
                return false;
            }
        }
        public override bool Equals(object obj)
        {
            return this.Equals((XLColor)obj);
        }

        int hashCode;
        public override int GetHashCode()
        {
            return hashCode;
        }
    }
}
