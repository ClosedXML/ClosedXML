using System;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

internal sealed partial class XLBorderCellFormat : IXLBorder
{
    XLBorderStyleValues IXLBorder.OutsideBorder
    {
        set => _parent.ModifyOuterBorder(static (borderLine, style) => borderLine with { Style = style }, value);
    }

    XLColor IXLBorder.OutsideBorderColor
    {
        set => _parent.ModifyOuterBorder(static (borderLine, color) => borderLine with { Color = color }, value);
    }

    XLBorderStyleValues IXLBorder.InsideBorder
    {
        set => _parent.ModifyInnerBorder(static (borderLine, style) => borderLine with { Style = style }, value);
    }

    XLColor IXLBorder.InsideBorderColor
    {
        set => _parent.ModifyInnerBorder(static (borderLine, color) => borderLine with { Color = color }, value);
    }

    XLBorderStyleValues IXLBorder.LeftBorder
    {
        get => LeftBorder;
        set => LeftBorder = value;
    }

    XLColor IXLBorder.LeftBorderColor
    {
        get => LeftBorderColor;
        set => LeftBorderColor = value;
    }

    XLBorderStyleValues IXLBorder.RightBorder
    {
        get => RightBorder;
        set => RightBorder = value;
    }

    XLColor IXLBorder.RightBorderColor
    {
        get => RightBorderColor;
        set => RightBorderColor = value;
    }

    XLBorderStyleValues IXLBorder.TopBorder
    {
        get => TopBorder;
        set => TopBorder = value;
    }

    XLColor IXLBorder.TopBorderColor
    {
        get => TopBorderColor;
        set => TopBorderColor = value;
    }

    XLBorderStyleValues IXLBorder.BottomBorder
    {
        get => BottomBorder;
        set => BottomBorder = value;
    }

    XLColor IXLBorder.BottomBorderColor
    {
        get => BottomBorderColor;
        set => BottomBorderColor = value;
    }

    bool IXLBorder.DiagonalUp
    {
        get => DiagonalUp;
        set => DiagonalUp = value;
    }

    bool IXLBorder.DiagonalDown
    {
        get => DiagonalDown;
        set => DiagonalDown = value;
    }

    XLBorderStyleValues IXLBorder.DiagonalBorder
    {
        get => DiagonalBorder;
        set => DiagonalBorder = value;
    }

    XLColor IXLBorder.DiagonalBorderColor
    {
        get => DiagonalBorderColor;
        set => DiagonalBorderColor = value;
    }

    IXLStyle IXLBorder.SetOutsideBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).OutsideBorder = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetOutsideBorderColor(XLColor value)
    {
        (this as IXLBorder).OutsideBorderColor = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetInsideBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).InsideBorder = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetInsideBorderColor(XLColor value)
    {
        (this as IXLBorder).InsideBorderColor = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetLeftBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).LeftBorder = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetLeftBorderColor(XLColor value)
    {
        (this as IXLBorder).LeftBorderColor = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetRightBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).RightBorder = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetRightBorderColor(XLColor value)
    {
        (this as IXLBorder).RightBorderColor = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetTopBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).TopBorder = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetTopBorderColor(XLColor value)
    {
        (this as IXLBorder).TopBorderColor = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetBottomBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).BottomBorder = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetBottomBorderColor(XLColor value)
    {
        (this as IXLBorder).BottomBorderColor = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetDiagonalUp()
    {
        return (this as IXLBorder).SetDiagonalUp(true);
    }

    IXLStyle IXLBorder.SetDiagonalUp(bool value)
    {
        (this as IXLBorder).DiagonalUp = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetDiagonalDown()
    {
        return (this as IXLBorder).SetDiagonalDown(true);
    }

    IXLStyle IXLBorder.SetDiagonalDown(bool value)
    {
        (this as IXLBorder).DiagonalDown = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetDiagonalBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).DiagonalBorder = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetDiagonalBorderColor(XLColor value)
    {
        (this as IXLBorder).DiagonalBorderColor = value;
        return _parent;
    }

    bool IEquatable<IXLBorder>.Equals(IXLBorder? other)
    {
        // A "business" equals, when borders are visually same, they are considered equals.
        if (other is null)
            return false;

        var thisBorder = _parent.Resolve(static x => x.Border);
        var otherLeft = new XLBorderLine(other.LeftBorderColor, other.LeftBorder);
        if (!IsSameLine(thisBorder.Left, otherLeft))
            return false;

        var otherTop = new XLBorderLine(other.TopBorderColor, other.TopBorder);
        if (!IsSameLine(thisBorder.Top, otherTop))
            return false;

        var otherRight = new XLBorderLine(other.RightBorderColor, other.RightBorder);
        if (!IsSameLine(thisBorder.Right, otherRight))
            return false;

        var otherBottom = new XLBorderLine(other.BottomBorderColor, other.BottomBorder);
        if (!IsSameLine(thisBorder.Bottom, otherBottom))
            return false;

        // Check diagonals. If the diagonal flag is not set, the diagonal is not displayed. Normalize them to take into the account the direction flag
        var thisDiagonalUp = MakeThisDiagonal(thisBorder.Diagonal, thisBorder.DiagonalUp);
        var otherDiagonalUp = MakeOtherDiagonal(other, other.DiagonalUp);
        if (!IsSameLine(thisDiagonalUp, otherDiagonalUp))
            return false;

        var thisDiagonalDown = MakeThisDiagonal(thisBorder.Diagonal, thisBorder.DiagonalDown);
        var otherDiagonalDown = MakeOtherDiagonal(other, other.DiagonalDown);
        if (!IsSameLine(thisDiagonalDown, otherDiagonalDown))
            return false;

        return true;

        static bool IsSameLine(XLBorderLine lhs, XLBorderLine rhs)
        {
            if (lhs.Style == XLBorderStyleValues.None && rhs.Style == XLBorderStyleValues.None)
                return true;

            // Auto color in context of border is black, not transparent.
            return lhs.Style == rhs.Style && lhs.Color == rhs.Color;
        }

        static XLBorderLine MakeThisDiagonal(XLBorderLine diagonal, bool diagonalDirection)
        {
            return diagonal with { Style = diagonalDirection ? diagonal.Style : XLBorderStyleValues.None };
        }

        static XLBorderLine MakeOtherDiagonal(IXLBorder border, bool diagonalDirection)
        {
            return new XLBorderLine(border.DiagonalBorderColor, diagonalDirection ? border.DiagonalBorder : XLBorderStyleValues.None);
        }
    }
}
