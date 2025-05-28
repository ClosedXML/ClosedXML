namespace ClosedXML.Excel.Formatting;

internal record XLAlignmentFormatValue
{
    public required XLAlignmentHorizontalValues Horizontal { get; init; }

    public required XLAlignmentVerticalValues Vertical { get; init; }

    public required TextRotation TextRotation { get; init; }

    /// <summary>
    /// Should the text be line-wrapped within the cell?
    /// </summary>
    public required bool WrapText { get; init; }

    /// <summary>
    /// Indicates number of 3*spaces (of the normal style font) of indentation for text in a cell.
    /// </summary>
    public required int Indent { get; init; }

    /// <summary>
    /// Relative indentation for dxf. Indicates number of spaces to indent text in a cell.
    /// </summary>
    public required int RelativeIndent { get; init; }

    public required bool JustifyLastLine { get; init; }

    public required bool ShrinkToFit { get; init; }

    public required XLAlignmentReadingOrderValues ReadingOrder { get; init; }

    internal XLAlignmentKey ApplyTo(XLAlignmentKey alignmentKey)
    {
        return new XLAlignmentKey
        {
            Horizontal = Horizontal,
            Vertical = Vertical,
            TextRotation = TextRotation.Value,
            WrapText = WrapText,
            Indent = Indent,
            RelativeIndent = RelativeIndent,
            JustifyLastLine = JustifyLastLine,
            ShrinkToFit = ShrinkToFit,
            ReadingOrder = ReadingOrder
        };
    }
}
