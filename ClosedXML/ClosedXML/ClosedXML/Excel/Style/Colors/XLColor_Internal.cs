using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ClosedXML.Excel
{
    public partial class XLColor
    {
        
        internal XLColor(IXLColor defaultColor)
        {
            if (defaultColor.ColorType == XLColorType.Theme)
            {
                this.ThemeColor = defaultColor.ThemeColor;
                this.ThemeTint = defaultColor.ThemeTint;
                hashCode = 7 ^ themeColor.GetHashCode() ^ themeTint.GetHashCode();
            }
            else if (defaultColor.ColorType == XLColorType.Indexed)
            {
                this.Indexed = defaultColor.Indexed;
                hashCode = 11 ^ indexed;
            }
            else
            {
                this.Color = defaultColor.Color;
                hashCode = 13 ^ color.GetHashCode();
            }
            HasValue = true;
        }
        internal XLColor()
        {
            HasValue = false;
            hashCode = 0;
        }
        internal XLColor(Color color)
        {
            Color = color;
            hashCode = 13 ^ this.color.GetHashCode();
            HasValue = true;
        }
        internal XLColor(Int32 index)
        {
            Indexed = index;
            hashCode = 11 ^ indexed;
            HasValue = true;
        }
        internal XLColor(XLThemeColor themeColor)
        {
            ThemeColor = themeColor;
            ThemeTint = 1;
            hashCode = 7 ^ this.themeColor.GetHashCode() ^ themeTint.GetHashCode();
            HasValue = true;
        }
        internal XLColor(XLThemeColor themeColor, Double themeTint)
        {
            ThemeColor = themeColor;
            ThemeTint = themeTint;
            hashCode = 7 ^ this.themeColor.GetHashCode() ^ this.themeTint.GetHashCode();
            HasValue = true;
        }
    }
}
