using System;
using System.Drawing;
using System.Globalization;

namespace ClosedXML.Excel
{
    public enum XLColorType
    {
        /// <summary>
        /// Automatic color. The actual color is determined by the application depending on a context where it is used.
        /// Generally speaking, the value is resolved either as a black (e.g. border or font color) or as a white (cell
        /// or chart fill). The <see cref="XLColor.Color"/> of automatic color has no bearing on actual resolved color
        /// and should be ignored.
        /// </summary>
        Automatic,

        /// <summary>
        /// A RGB color. It can technically specify alpha component, but Excel just ignores that and marks everything
        /// as fully opaque. The color value is stored directly in <see cref="XLColor.Color"/>.
        /// </summary>
        Color,

        /// <summary>
        /// A theme color. The color value depends on a theme of a workbook and can be resolved through <see cref="IXLTheme.ResolveThemeColor(XLThemeColor)"/>.
        /// </summary>
        Theme,

        /// <summary>
        /// An indexed color. Only for legacy usage, used in times when palette was common. The only semi-valid usage
        /// is for system foreground color (index 64) and system background color (index 65). The default indexed colors can
        /// be found in <see cref="XLColor.Indexed"/> and the <see cref="XLColor.Color"/> will return a value that
        /// corresponds to the default indexed color.
        /// </summary>
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
        internal Boolean IsAuto => !HasValue;

        public Boolean HasValue { get; }

        public XLColorType ColorType => Key.ColorType;

        public Color Color
        {
            get
            {
                if (ColorType == XLColorType.Color)
                    return Key.Color;

                if (ColorType == XLColorType.Indexed)
                    return IndexedColors[Indexed].Color;

                throw new InvalidOperationException($"Cannot convert {LcColorType} color to Color.");
            }
        }

        public Int32 Indexed
        {
            get
            {
                if (ColorType == XLColorType.Indexed)
                    return Key.Indexed;

                throw new InvalidOperationException($"Cannot convert {LcColorType} color to indexed color.");

            }
        }

        public XLThemeColor ThemeColor
        {
            get
            {
                if (ColorType == XLColorType.Theme)
                    return Key.ThemeColor;

                throw new InvalidOperationException($"Cannot convert {LcColorType} color to theme color.");
            }
        }

        public Double ThemeTint
        {
            get
            {
                if (ColorType == XLColorType.Theme)
                    return Key.ThemeTint;

                if (ColorType == XLColorType.Indexed)
                    throw new InvalidOperationException("Cannot extract theme tint from an indexed color.");

                return Color.A / 255.0;
            }
        }

        #region IEquatable<XLColor> Members

        public bool Equals(XLColor other)
        {
            return Key == other.Key;
        }

        #endregion IEquatable<XLColor> Members

        public override bool Equals(object obj)
        {
            return Equals((XLColor)obj);
        }

        public override int GetHashCode()
        {
            var hashCode = 229333804;
            hashCode = hashCode * -1521134295 + HasValue.GetHashCode();
            hashCode = hashCode * -1521134295 + Key.GetHashCode();
            return hashCode;
        }

        public override string ToString()
        {
            if (ColorType == XLColorType.Color)
                return Color.ToHex();

            if (ColorType == XLColorType.Theme)
                return $"Color Theme: {ThemeColor}, Tint: {ThemeTint.ToString(CultureInfo.InvariantCulture)}";

            if (ColorType == XLColorType.Automatic)
                return "Automatic";

            return "Color Index: " + Indexed;
        }

        public static Boolean operator ==(XLColor? left, XLColor? right)
        {
            // If both are null, or both are same instance, return true.
            if (ReferenceEquals(left, right)) return true;

            // If one is null, but not both, return false.
            if ((left as object) == null || (right as Object) == null) return false;

            return left.Equals(right);
        }

        public static Boolean operator !=(XLColor? left, XLColor? right)
        {
            return !(left == right);
        }
    }
}
