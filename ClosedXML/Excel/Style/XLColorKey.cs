#nullable disable
using System;

namespace ClosedXML.Excel;

internal struct XLColorKey : IEquatable<XLColorKey>
{
    public XLColorType ColorType { get; set; }

    public System.Drawing.Color Color { get; set; }

    public int Indexed { get; set; }

    public XLThemeColor ThemeColor { get; set; }

    public double ThemeTint { get; set; }

    public override int GetHashCode()
    {
        unchecked
        {
            var hash = (int)ColorType;

            switch (ColorType)
            {
                case XLColorType.Indexed:
                    hash = (hash * 397) ^ Indexed;
                    break;

                case XLColorType.Theme:
                    hash = (hash * 397) ^ (int)ThemeColor;

                    // Normalize ThemeTint so that small float differences don't cause large hash differences
                    var tintHash = (int)(ThemeTint * 100000); // Multiply to make epsilon differences visible
                    hash = (hash * 397) ^ tintHash;
                    break;

                case XLColorType.Color:
                    hash = (hash * 397) ^ Color.ToArgb();
                    break;
            }

            return hash;
        }
    }

    public bool Equals(XLColorKey other)
    {
        if (ColorType == other.ColorType)
        {
            if (ColorType == XLColorType.Color)
            {
                // .NET Color.Equals() will return false for Color.FromArgb(255, 255, 255, 255) == Color.White
                // Therefore we compare the ToArgb() values
                return Color.ToArgb() == other.Color.ToArgb();
            }
            if (ColorType == XLColorType.Theme)
            {
                return
                    ThemeColor == other.ThemeColor
                    && Math.Abs(ThemeTint - other.ThemeTint) < XLHelper.Epsilon;
            }
            return Indexed == other.Indexed;
        }

        return false;
    }

    public override bool Equals(object obj)
    {
        if (obj is XLColorKey key)
            return Equals(key);
        return base.Equals(obj);
    }

    public override string ToString()
    {
        return ColorType switch
        {
            XLColorType.Color => Color.ToString(),
            XLColorType.Theme => $"{ThemeColor} ({ThemeTint})",
            XLColorType.Indexed => $"Indexed: {Indexed}",
            _ => base.ToString()
        };
    }

    public static bool operator ==(XLColorKey left, XLColorKey right) => left.Equals(right);

    public static bool operator !=(XLColorKey left, XLColorKey right) => !(left.Equals(right));
}