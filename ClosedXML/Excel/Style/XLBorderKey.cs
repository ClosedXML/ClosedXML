namespace ClosedXML.Excel;

internal readonly record struct XLBorderKey
{
    public required XLBorderStyleValues LeftBorder { get; init; }

    public required XLColorKey LeftBorderColor { get; init; }

    public required XLBorderStyleValues RightBorder { get; init; }

    public required XLColorKey RightBorderColor { get; init; }

    public required XLBorderStyleValues TopBorder { get; init; }

    public required XLColorKey TopBorderColor { get; init; }

    public required XLBorderStyleValues BottomBorder { get; init; }

    public required XLColorKey BottomBorderColor { get; init; }

    public required XLBorderStyleValues DiagonalBorder { get; init; }

    public required XLColorKey DiagonalBorderColor { get; init; }

    public required bool DiagonalUp { get; init; }

    public required bool DiagonalDown { get; init; }

    public override int GetHashCode()
    {
        unchecked
        {
            var hash = (int)LeftBorder;
            hash = (hash * 397) ^ (LeftBorder != XLBorderStyleValues.None ? LeftBorderColor.GetHashCode() : 0);

            hash = (hash * 397) ^ (int)RightBorder;
            hash = (hash * 397) ^ (RightBorder != XLBorderStyleValues.None ? RightBorderColor.GetHashCode() : 0);

            hash = (hash * 397) ^ (int)TopBorder;
            hash = (hash * 397) ^ (TopBorder != XLBorderStyleValues.None ? TopBorderColor.GetHashCode() : 0);

            hash = (hash * 397) ^ (int)BottomBorder;
            hash = (hash * 397) ^ (BottomBorder != XLBorderStyleValues.None ? BottomBorderColor.GetHashCode() : 0);

            hash = (hash * 397) ^ (int)DiagonalBorder;
            hash = (hash * 397) ^ (DiagonalBorder != XLBorderStyleValues.None ? DiagonalBorderColor.GetHashCode() : 0);

            hash = (hash * 397) ^ (DiagonalUp ? 1 : 0);
            hash = (hash * 397) ^ (DiagonalDown ? 1 : 0);

            return hash;
        }
    }

    public bool Equals(XLBorderKey other)
    {
        return
            AreEquivalent(LeftBorder, LeftBorderColor, other.LeftBorder, other.LeftBorderColor)
            && AreEquivalent(RightBorder, RightBorderColor, other.RightBorder, other.RightBorderColor)
            && AreEquivalent(TopBorder, TopBorderColor, other.TopBorder, other.TopBorderColor)
            && AreEquivalent(BottomBorder, BottomBorderColor, other.BottomBorder, other.BottomBorderColor)
            && AreEquivalent(DiagonalBorder, DiagonalBorderColor, other.DiagonalBorder, other.DiagonalBorderColor)
            && DiagonalUp == other.DiagonalUp
            && DiagonalDown == other.DiagonalDown;
    }

    private static bool AreEquivalent(
        XLBorderStyleValues borderStyle1, XLColorKey color1,
        XLBorderStyleValues borderStyle2, XLColorKey color2)
    {
        return (borderStyle1 == XLBorderStyleValues.None &&
                borderStyle2 == XLBorderStyleValues.None) ||
               borderStyle1 == borderStyle2 &&
               color1 == color2;
    }

    public override string ToString()
    {
        return $"{LeftBorder} {LeftBorderColor} {RightBorder} {RightBorderColor} {TopBorder} {TopBorderColor} " +
               $"{BottomBorder} {BottomBorderColor} {DiagonalBorder} {DiagonalBorderColor} " +
               (DiagonalUp ? "DiagonalUp" : "") +
               (DiagonalDown ? "DiagonalDown" : "");
    }
}