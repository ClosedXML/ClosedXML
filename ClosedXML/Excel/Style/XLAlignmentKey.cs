namespace ClosedXML.Excel;

public readonly record struct XLAlignmentKey
{
    public required XLAlignmentHorizontalValues Horizontal { get; init; }

    public required XLAlignmentVerticalValues Vertical { get; init; }

    public required int Indent { get; init; }

    public required bool JustifyLastLine { get; init; }

    public required XLAlignmentReadingOrderValues ReadingOrder { get; init; }

    public required int RelativeIndent { get; init; }

    public required bool ShrinkToFit { get; init; }

    public required int TextRotation { get; init; }

    public required bool WrapText { get; init; }

    public override string ToString()
    {
        return
            $"{Horizontal} {Vertical} {ReadingOrder} Indent: {Indent} RelativeIndent: {RelativeIndent} TextRotation: {TextRotation} " +
            (WrapText ? "WrapText" : "") +
            (JustifyLastLine ? "JustifyLastLine" : "");
    }
}
