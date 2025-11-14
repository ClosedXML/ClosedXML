using System;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// Explicit implementation of the <see cref="IXLAlignment"/> interface.
/// </summary>
internal sealed partial class XLAlignmentCellFormat : IXLAlignment
{
    XLAlignmentHorizontalValues IXLAlignment.Horizontal
    {
        get => Resolve(static format => format.Alignment.Horizontal);
        set => Modify(static (alignment, hAlign) => alignment with { Horizontal = hAlign }, value);
    }

    XLAlignmentVerticalValues IXLAlignment.Vertical
    {
        get => Resolve(static format => format.Alignment.Vertical);
        set => Modify(static (alignment, vAlign) => alignment with { Vertical = vAlign }, value);
    }

    int IXLAlignment.Indent
    {
        get => Resolve(static format => format.Alignment.Indent);
        set => Modify(static (alignment, indent) => alignment with { Indent = indent }, value);
    }

    bool IXLAlignment.JustifyLastLine
    {
        get => Resolve(static format => format.Alignment.JustifyLastLine);
        set => Modify(static (alignment, justifyLastLine) => alignment with { JustifyLastLine = justifyLastLine }, value);
    }

    XLAlignmentReadingOrderValues IXLAlignment.ReadingOrder
    {
        get => Resolve(static format => format.Alignment.ReadingOrder);
        set => Modify(static (alignment, readingOrder) => alignment with { ReadingOrder = readingOrder }, value);
    }

    int IXLAlignment.RelativeIndent
    {
        get => Resolve(static format => format.Alignment.RelativeIndent);
        set => Modify(static (alignment, relativeIndent) => alignment with { RelativeIndent = relativeIndent }, value);
    }

    bool IXLAlignment.ShrinkToFit
    {
        get => Resolve(static format => format.Alignment.ShrinkToFit);
        set => Modify(static (alignment, shrinkToFit) => alignment with { ShrinkToFit = shrinkToFit }, value);
    }

    int IXLAlignment.TextRotation
    {
        get => Resolve(static format => format.Alignment.TextRotation.Value);
        set => Modify(static (alignment, textRotation) => alignment with { TextRotation = new TextRotation(textRotation) }, value);
    }

    bool IXLAlignment.WrapText
    {
        get => Resolve(static format => format.Alignment.WrapText);
        set => Modify(static (alignment, wrapText) => alignment with { WrapText = wrapText }, value);
    }

    bool IXLAlignment.TopToBottom
    {
        get => Resolve(static format => format.Alignment.TextRotation == TextRotation.VerticalText);
        set => Modify(static (alignment, topToBottom) => alignment with { TextRotation = topToBottom ? TextRotation.VerticalText : TextRotation.None }, value);
    }

    IXLStyle IXLAlignment.SetHorizontal(XLAlignmentHorizontalValues value)
    {
        (this as IXLAlignment).Horizontal = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetVertical(XLAlignmentVerticalValues value)
    {
        (this as IXLAlignment).Vertical = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetIndent(int value)
    {
        (this as IXLAlignment).Indent = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetJustifyLastLine()
    {
        return (this as IXLAlignment).SetJustifyLastLine(true);
    }

    IXLStyle IXLAlignment.SetJustifyLastLine(bool value)
    {
        (this as IXLAlignment).JustifyLastLine = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetReadingOrder(XLAlignmentReadingOrderValues value)
    {
        (this as IXLAlignment).ReadingOrder = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetRelativeIndent(int value)
    {
        (this as IXLAlignment).RelativeIndent = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetShrinkToFit()
    {
        return (this as IXLAlignment).SetShrinkToFit(true);
    }

    IXLStyle IXLAlignment.SetShrinkToFit(bool value)
    {
        (this as IXLAlignment).ShrinkToFit = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetTextRotation(int value)
    {
        (this as IXLAlignment).TextRotation = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetWrapText()
    {
        return (this as IXLAlignment).SetWrapText(true);
    }

    IXLStyle IXLAlignment.SetWrapText(bool value)
    {
        (this as IXLAlignment).WrapText = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetTopToBottom()
    {
        return (this as IXLAlignment).SetTopToBottom(true);
    }

    IXLStyle IXLAlignment.SetTopToBottom(bool value)
    {
        (this as IXLAlignment).TopToBottom = value;
        return _parent;
    }

    bool IEquatable<IXLAlignment>.Equals(IXLAlignment? other)
    {
        if (other == null)
            return false;

        var align = Resolve(static x => x.Alignment);
        if (align.Horizontal != other.Horizontal)
            return false;

        if (align.Vertical != other.Vertical)
            return false;

        if (align.Indent != other.Indent)
            return false;

        if (align.JustifyLastLine != other.JustifyLastLine)
            return false;

        if (align.ReadingOrder != other.ReadingOrder)
            return false;

        if (align.RelativeIndent != other.RelativeIndent)
            return false;

        if (align.ShrinkToFit != other.ShrinkToFit)
            return false;

        if (align.TextRotation.Value != other.TextRotation)
            return false;

        if (align.WrapText != other.WrapText)
            return false;

        return true;
    }
}
