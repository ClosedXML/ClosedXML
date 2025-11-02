using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

internal sealed partial class XLBorderCellFormat
{
    private readonly XLCellFormat _parent;

    internal XLBorderCellFormat(XLCellFormat parent)
    {
        _parent = parent;
    }

    internal XLBorderStyleValues LeftBorder
    {
        get => _parent.Resolve(static x => x.Border.Left.Style);
        set => _parent.ModifyBorder(static (border, leftStyle) => border with { Left = border.Left with { Style = leftStyle } }, value);
    }

    internal XLColor LeftBorderColor
    {
        get => _parent.Resolve(static x => x.Border.Left.Color);
        set => _parent.ModifyBorder(static (border, leftColor) => border with { Left = border.Left with { Color = leftColor } }, value);
    }

    internal XLBorderStyleValues RightBorder
    {
        get => _parent.Resolve(static x => x.Border.Right.Style);
        set => _parent.ModifyBorder(static (border, rightStyle) => border with { Right = border.Right with { Style = rightStyle } }, value);
    }

    internal XLColor RightBorderColor
    {
        get => _parent.Resolve(static x => x.Border.Right.Color);
        set => _parent.ModifyBorder(static (border, rightColor) => border with { Right = border.Right with { Color = rightColor } }, value);
    }

    internal XLBorderStyleValues TopBorder
    {
        get => _parent.Resolve(static x => x.Border.Top.Style);
        set => _parent.ModifyBorder(static (border, topStyle) => border with { Top = border.Top with { Style = topStyle } }, value);
    }

    internal XLColor TopBorderColor
    {
        get => _parent.Resolve(static x => x.Border.Top.Color);
        set => _parent.ModifyBorder(static (border, topColor) => border with { Top = border.Top with { Color = topColor } }, value);
    }

    internal XLBorderStyleValues BottomBorder
    {
        get => _parent.Resolve(static x => x.Border.Bottom.Style);
        set => _parent.ModifyBorder(static (border, bottomStyle) => border with { Bottom = border.Bottom with { Style = bottomStyle } }, value);
    }

    internal XLColor BottomBorderColor
    {
        get => _parent.Resolve(static x => x.Border.Bottom.Color);
        set => _parent.ModifyBorder(static (border, bottomColor) => border with { Bottom = border.Bottom with { Color = bottomColor } }, value);
    }

    internal bool DiagonalUp
    {
        get => _parent.Resolve(static x => x.Border.DiagonalUp);
        set => _parent.ModifyBorder(static (border, diagonalUp) => border with { DiagonalUp = diagonalUp }, value);
    }

    internal bool DiagonalDown
    {
        get => _parent.Resolve(static x => x.Border.DiagonalDown);
        set => _parent.ModifyBorder(static (border, diagonalDown) => border with { DiagonalDown = diagonalDown }, value);
    }

    internal XLBorderStyleValues DiagonalBorder
    {
        get => _parent.Resolve(static x => x.Border.Diagonal.Style);
        set => _parent.ModifyBorder(static (border, diagonalStyle) => border with { Diagonal = border.Diagonal with { Style = diagonalStyle } }, value);
    }

    internal XLColor DiagonalBorderColor
    {
        get => _parent.Resolve(static x => x.Border.Diagonal.Color);
        set => _parent.ModifyBorder(static (border, diagonalColor) => border with { Diagonal = border.Diagonal with { Color = diagonalColor } }, value);
    }

    internal void SetValue(IXLBorder value)
    {
        _parent.ModifyBorder((border, other) => border with
        {
            Left = new XLBorderLine(other.LeftBorderColor, other.LeftBorder),
            Right = new XLBorderLine(other.RightBorderColor, other.RightBorder),
            Top = new XLBorderLine(other.TopBorderColor, other.TopBorder),
            Bottom = new XLBorderLine(other.BottomBorderColor, other.BottomBorder),
            DiagonalUp = other.DiagonalUp,
            DiagonalDown = other.DiagonalDown,
            Diagonal = new XLBorderLine(other.DiagonalBorderColor, other.DiagonalBorder),
        }, value);
    }
}
