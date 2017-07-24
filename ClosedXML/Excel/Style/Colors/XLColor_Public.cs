using System;
using System.Drawing;

namespace ClosedXML.Excel
{
    public enum XLColorType
    {
        Color,
        Theme,
        Indexed
    }

    public enum XLThemeColor
    {
        Background1,
        Text1,
        Background2,
        Text2,
        Accent1,
        Accent2,
        Accent3,
        Accent4,
        Accent5,
        Accent6,
        Hyperlink,
        FollowedHyperlink
    }

    public partial class XLColor : IEquatable<XLColor>
    {
        /// <summary>
        /// Usually indexed colors are limited to max 63
        /// Index 81 is some special case.
        /// Some people claim it's the index for tooltip color.
        /// We'll return normal black when index 81 is found.
        /// </summary>
        private const Int32 TOOLTIPCOLORINDEX = 81;

        private readonly XLColorType _colorType;
        private int _hashCode;
        private readonly Int32 _indexed;
        private readonly XLThemeColor _themeColor;
        private readonly Double _themeTint;

        private Color _color;
        public Boolean HasValue { get; private set; }

        public XLColorType ColorType
        {
            get { return _colorType; }
        }

        public Color Color
        {
            get
            {
                if (_colorType == XLColorType.Theme)
                    throw new Exception("Cannot convert theme color to Color.");

                if (_colorType == XLColorType.Indexed)
                    if (_indexed == TOOLTIPCOLORINDEX)
                        return Color.FromArgb(255, Color.Black);
                    else
                        return IndexedColors[_indexed].Color;

                return _color;
            }
        }

        public Int32 Indexed
        {
            get
            {
                if (ColorType == XLColorType.Theme)
                    throw new Exception("Cannot convert theme color to indexed color.");

                if (ColorType == XLColorType.Indexed)
                    return _indexed;

                throw new Exception("Cannot convert Color to indexed color.");
            }
        }

        public XLThemeColor ThemeColor
        {
            get
            {
                if (ColorType == XLColorType.Theme)
                    return _themeColor;

                if (ColorType == XLColorType.Indexed)
                    throw new Exception("Cannot convert indexed color to theme color.");

                throw new Exception("Cannot convert Color to theme color.");
            }
        }

        public Double ThemeTint
        {
            get
            {
                if (ColorType == XLColorType.Theme)
                    return _themeTint;

                if (ColorType == XLColorType.Indexed)
                    throw new Exception("Cannot extract theme tint from an indexed color.");

                return _color.A/255.0;
            }
        }

        #region IEquatable<XLColor> Members

        public bool Equals(XLColor other)
        {
            if (_colorType == other._colorType)
            {
                if (_colorType == XLColorType.Color)
                {
                    // .NET Color.Equals() will return false for Color.FromArgb(255, 255, 255, 255) == Color.White
                    // Therefore we compare the ToArgb() values
                    return _color.ToArgb() == other._color.ToArgb();
                }
                if (_colorType == XLColorType.Theme)
                {
                    return _themeColor == other._themeColor
                           && Math.Abs(_themeTint - other._themeTint) < XLHelper.Epsilon;
                }
                return _indexed == other._indexed;
            }

            return false;
        }

        #endregion

        public override bool Equals(object obj)
        {
            return Equals((XLColor) obj);
        }

        public override int GetHashCode()
        {
            if (_hashCode == 0)
            {
                if (_colorType == XLColorType.Color)
                    _hashCode = _color.GetHashCode();
                else if (_colorType == XLColorType.Theme)
                    _hashCode = _themeColor.GetHashCode() ^ _themeTint.GetHashCode();
                else
                    _hashCode = _indexed;
            }

            return _hashCode;
        }

        public override string ToString()
        {
            if (_colorType == XLColorType.Color)
                return Color.ToHex();

            if (_colorType == XLColorType.Theme)
                return String.Format("Color Theme: {0}, Tint: {1}", _themeColor.ToString(), _themeTint.ToString());

            return "Color Index: " + _indexed;
        }

        public static Boolean operator ==(XLColor left, XLColor right)
        {
            // If both are null, or both are same instance, return true.
            if (ReferenceEquals(left, right)) return true;

            // If one is null, but not both, return false.
            if ((left as object) == null || (right as Object) == null) return false;

            return left.Equals(right);
        }

        public static Boolean operator !=(XLColor left, XLColor right)
        {
            return !(left == right);
        }
    }
}
