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
            }
            else if (defaultColor.ColorType == XLColorType.Indexed)
            {
                this.Indexed = defaultColor.Indexed;
            }
            else
            {
                this.Color = defaultColor.Color;
            }
        }
        internal XLColor()
        {
            HasValue = false;
        }
        internal XLColor(Color color)
        {
            Color = color;
            HasValue = true;
        }
        internal XLColor(Int32 index)
        {
            Indexed = index;
            HasValue = true;
        }
        internal XLColor(XLThemeColor themeColor)
        {
            ThemeColor = themeColor;
            ThemeTint = 1;
            HasValue = true;
        }
        internal XLColor(XLThemeColor themeColor, Double themeTint)
        {
            ThemeColor = themeColor;
            ThemeTint = themeTint;
            HasValue = true;
        }
    }
}
