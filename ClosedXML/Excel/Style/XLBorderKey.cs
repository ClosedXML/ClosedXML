using System;

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
        var hash = new HashCode();
        hash.Add(LeftBorder);
        hash.Add(RightBorder);
        hash.Add(TopBorder);
        hash.Add(BottomBorder);
        hash.Add(DiagonalBorder);
        hash.Add(DiagonalUp);
        hash.Add(DiagonalDown);

        if (LeftBorder != XLBorderStyleValues.None)
            hash.Add(LeftBorderColor);
        if (RightBorder != XLBorderStyleValues.None)
            hash.Add(RightBorderColor);
        if (TopBorder != XLBorderStyleValues.None)
            hash.Add(TopBorderColor);
        if (BottomBorder != XLBorderStyleValues.None)
            hash.Add(BottomBorderColor);
        if (DiagonalBorder != XLBorderStyleValues.None)
            hash.Add(DiagonalBorderColor);

        return hash.ToHashCode();
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

    private bool AreEquivalent(
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
