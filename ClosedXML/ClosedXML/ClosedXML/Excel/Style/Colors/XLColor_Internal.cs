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
            XLColor dColor = (XLColor)defaultColor;
            if (dColor.colorType == XLColorType.Theme)
            {
                this.themeColor = dColor.themeColor;
                this.themeTint = dColor.themeTint;
                hashCode = 7 ^ themeColor.GetHashCode() ^ themeTint.GetHashCode();
            }
            else if (dColor.ColorType == XLColorType.Indexed)
            {
                this.indexed = dColor.indexed;
                hashCode = 11 ^ indexed;
            }
            else
            {
                this.color = dColor.color;
                hashCode = 13 ^ color.GetHashCode();
            }
            HasValue = true;
            colorType = dColor.colorType;
        }
        internal XLColor()
        {
            HasValue = false;
            hashCode = 0;
        }
        internal XLColor(Color color)
        {
            this.color = color;
            hashCode = 13 ^ this.color.GetHashCode();
            HasValue = true;
            colorType = XLColorType.Color;
        }
        internal XLColor(Int32 index)
        {
            this.indexed = index;
            hashCode = 11 ^ indexed;
            HasValue = true;
            colorType = XLColorType.Indexed;
        }
        internal XLColor(XLThemeColor themeColor)
        {
            this.themeColor = themeColor;
            this.themeTint = 1;
            hashCode = 7 ^ this.themeColor.GetHashCode() ^ themeTint.GetHashCode();
            HasValue = true;
            colorType = XLColorType.Theme;
        }
        internal XLColor(XLThemeColor themeColor, Double themeTint)
        {
            this.themeColor = themeColor;
            this.themeTint = themeTint;
            hashCode = 7 ^ this.themeColor.GetHashCode() ^ this.themeTint.GetHashCode();
            HasValue = true;
            colorType = XLColorType.Theme;
        }
    }
}
