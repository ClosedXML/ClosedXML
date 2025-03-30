namespace ClosedXML.Excel.Formatting;

/// <summary>
/// A formatting record for <see cref="XLCellFormat"/>. Unlike <see cref="XLFontKey"/>, attributes are optional.
/// </summary>
internal readonly record struct XLFontFormat
{
    public required XLFontName? Name { get; init; }

    public required XLFontCharSet? Charset { get; init; }

    public required XLFontFamilyNumberingValues? Family { get; init; }

    public required bool? Bold { get; init; }

    public required bool? Italic { get; init; }

    public required bool? Strikethrough { get; init; }

    public required bool? Outline { get; init; }

    public required bool? Shadow { get; init; }

    public required bool? Condense { get; init; }

    public required bool? Extend { get; init; }

    public required XLColor? Color { get; init; }

    public required XLFontSize? Size { get; init; }

    public required XLFontUnderlineValues? Underline { get; init; }

    public required XLFontVerticalTextAlignmentValues? VerticalAlignment { get; init; }

    public required XLFontScheme? Scheme { get; init; }
}
