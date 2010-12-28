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

        public XLColorType ColorType { get; private set; }
        private Color color;
        public Color Color 
        {
            get
            {
                if (ColorType == XLColorType.Theme)
                {
                    //if (workbook == null)
                        throw new Exception("Cannot convert theme color to Color.");
                    //else
                    //    return workbook.GetXLColor(themeColor).Color;
                }
                else if (ColorType == XLColorType.Indexed)
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
                ColorType = XLColorType.Color;
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
                ColorType = XLColorType.Indexed;
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
                ColorType = XLColorType.Theme;
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
                ColorType = XLColorType.Theme;
            }
        }

        public bool Equals(IXLColor other)
        {
            if (ColorType != other.ColorType)
            {
                return false;
            }
            else
            {
                if (ColorType == XLColorType.Theme)
                {
                    return this.ThemeColor.Equals(other.ThemeColor)
                        && this.ThemeTint.Equals(other.ThemeTint);
                }
                else if (ColorType == XLColorType.Indexed)
                {
                    return this.Indexed.Equals(other.Indexed);
                }
                else
                {
                    return this.Color.Equals(other.Color);
                }
            }
        }
        public override bool Equals(object obj)
        {
            return this.Equals((XLColor)obj);
        }

        public override int GetHashCode()
        {
            unchecked // Overflow is fine, just wrap
            {
                if (ColorType == XLColorType.Theme)
                {
                    int hash = 17;
                    hash = hash * 23 + ThemeColor.GetHashCode();
                    hash = hash * 23 + ThemeTint.GetHashCode();
                    return hash;
                }
                else if (ColorType == XLColorType.Indexed)
                {
                    return Indexed;
                }
                else
                {
                    return Color.GetHashCode();
                }
            }
        }
    }
}
