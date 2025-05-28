namespace ClosedXML.Excel.Formatting;

/// <summary>
/// A border format master record.
/// </summary>
internal record XLBorderFormatValue
{
    public required XLBorderLine? Left { get; init; }

    public required XLBorderLine? Right { get; init; }

    public required XLBorderLine? Top { get; init; }

    public required XLBorderLine? Bottom { get; init; }

    /// <summary>
    /// Used when <see cref="DiagonalUp"/> or <see cref="DiagonalDown"/> are set. It's not possible
    /// to have different style for up/down diagonal.
    /// </summary>
    public required XLBorderLine? Diagonal { get; init; }

    /// <summary>
    /// For pivot tables only.
    /// </summary>
    public required XLBorderLine? Vertical { get; init; }

    /// <summary>
    /// For pivot tables only.
    /// </summary>
    public required XLBorderLine? Horizontal { get; init; }

    public required bool DiagonalUp { get; init; }

    public required bool DiagonalDown { get; init; }

    public required bool Outline { get; init; }

    internal XLBorderKey ApplyTo(XLBorderKey borderKey)
    {
        if (Left is not null)
            borderKey = borderKey with { LeftBorder = Left.Value.Style, LeftBorderColor = Left.Value.Color.Key };

        if (Right is not null)
            borderKey = borderKey with { RightBorder = Right.Value.Style, RightBorderColor = Right.Value.Color.Key };

        if (Top is not null)
            borderKey = borderKey with { TopBorder = Top.Value.Style, TopBorderColor = Top.Value.Color.Key };

        if (Bottom is not null)
            borderKey = borderKey with { BottomBorder = Bottom.Value.Style, BottomBorderColor = Bottom.Value.Color.Key };

        if (Diagonal is not null)
        {
            borderKey = borderKey with
            {
                DiagonalBorder = Diagonal.Value.Style,
                DiagonalBorderColor = Diagonal.Value.Color.Key,
                DiagonalUp = DiagonalUp,
                DiagonalDown = DiagonalDown
            };
        }

        return borderKey;
    }
}
