using System;
using System.Drawing;

namespace ClosedXML.Excel
{
    public partial class XLColor
    {

        internal XLColor(IXLColor defaultColor)
        {
            XLColor dColor = (XLColor)defaultColor;
            if (dColor._colorType == XLColorType.Color)
                color = dColor.color;
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
            this.color = color;
            _hashCode = 13 ^ color.ToArgb();
            HasValue = true;
            _colorType = XLColorType.Color;
        }
        internal XLColor(Int32 index)
        {
            this._indexed = index;
            _hashCode = 11 ^ _indexed;
            HasValue = true;
            _colorType = XLColorType.Indexed;
        }
        internal XLColor(XLThemeColor themeColor)
        {
            this._themeColor = themeColor;
            this._themeTint = 1;
            _hashCode = 7 ^ this._themeColor.GetHashCode() ^ _themeTint.GetHashCode();
            HasValue = true;
            _colorType = XLColorType.Theme;
        }
        internal XLColor(XLThemeColor themeColor, Double themeTint)
        {
            this._themeColor = themeColor;
            this._themeTint = themeTint;
            _hashCode = 7 ^ this._themeColor.GetHashCode() ^ this._themeTint.GetHashCode();
            HasValue = true;
            _colorType = XLColorType.Theme;
        }
    }
}
