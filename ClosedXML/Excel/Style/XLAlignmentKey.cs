namespace ClosedXML.Excel;

public readonly record struct XLAlignmentKey
{
    public XLAlignmentHorizontalValues Horizontal { get; init; }

    public XLAlignmentVerticalValues Vertical { get; init; }

    public int Indent { get; init; }

    public bool JustifyLastLine { get; init; }

    public XLAlignmentReadingOrderValues ReadingOrder { get; init; }

    public int RelativeIndent { get; init; }

    public bool ShrinkToFit { get; init; }

    public int TextRotation { get; init; }

    public bool WrapText { get; init; }

    public bool TopToBottom { get; init; }

    public override string ToString()
    {
        return
            $"{Horizontal} {Vertical} {ReadingOrder} Indent: {Indent} RelativeIndent: {RelativeIndent} TextRotation: {TextRotation} " +
            (WrapText ? "WrapText" : "") +
            (JustifyLastLine ? "JustifyLastLine" : "") +
            (TopToBottom ? "TopToBottom" : "");
    }
}
