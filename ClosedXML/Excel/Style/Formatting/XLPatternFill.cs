namespace ClosedXML.Excel.Formatting;

/// <summary>
/// A pattern fill format.
/// </summary>
internal record XLPatternFill
{
    /// <summary>
    /// Foreground color of the pattern (i.e., if pattern is a grid, this is the color to draw
    /// lines of the grid).
    /// </summary>
    public required XLColor PatternColor { get; init; }

    /// <summary>
    /// Background color of the pattern.
    /// </summary>
    public required XLColor BackgroundColor { get; init; }

    /// <summary>
    /// Shape of the pattern.
    /// </summary>
    public required XLFillPatternValues PatternType { get; init; }
}
