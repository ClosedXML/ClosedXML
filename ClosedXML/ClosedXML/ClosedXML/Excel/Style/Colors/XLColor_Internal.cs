using System;
using System.Drawing;

namespace ClosedXML.Excel
{
    public partial class XLColor
    {
        internal XLColor(XLColor defaultColor)
        {
            var dColor = defaultColor;
            if (dColor._colorType == XLColorType.Color)
                _color = dColor._color;
            else if (dColor._colorType == XLColorType.Theme)
            {
                _themeColor = dColor._themeColor;
                _themeTint = dColor._themeTint;
            }
            else
            {
                _indexed = dColor._indexed;
            }

            HasValue = true;
            _hashCode = dColor._hashCode;
            _colorType = dColor._colorType;
        }

        internal XLColor()
        {
            HasValue = false;
            _hashCode = 0;
        }

        internal XLColor(Color color)
        {
            _color = color;
            _hashCode = 13 ^ color.ToArgb();
            HasValue = true;
            _colorType = XLColorType.Color;
        }

        internal XLColor(Int32 index)
        {
            _indexed = index;
            _hashCode = 11 ^ _indexed;
            HasValue = true;
            _colorType = XLColorType.Indexed;
        }

        internal XLColor(XLThemeColor themeColor)
        {
            _themeColor = themeColor;
            _themeTint = 1;
            _hashCode = 7 ^ _themeColor.GetHashCode() ^ _themeTint.GetHashCode();
            HasValue = true;
            _colorType = XLColorType.Theme;
        }

        internal XLColor(XLThemeColor themeColor, Double themeTint)
        {
            _themeColor = themeColor;
            _themeTint = themeTint;
            _hashCode = 7 ^ _themeColor.GetHashCode() ^ _themeTint.GetHashCode();
            HasValue = true;
            _colorType = XLColorType.Theme;
        }
    }
}