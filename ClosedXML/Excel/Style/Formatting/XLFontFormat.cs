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

    public XLFontKey ApplyTo(XLFontKey nf)
    {
        // No Outline, Condense or Extend
        if (Name is not null)
            nf = nf with { FontName = Name.Value.Text };

        if (Charset is not null)
            nf = nf with { FontCharSet = Charset.Value };

        if (Family is not null)
            nf = nf with { FontFamilyNumbering = Family.Value };

        if (Bold is not null)
            nf = nf with { Bold = Bold.Value };

        if (Italic is not null)
            nf = nf with { Italic = Italic.Value };

        if (Strikethrough is not null)
            nf = nf with { Strikethrough = Strikethrough.Value };

        if (Shadow is not null)
            nf = nf with { Shadow = Shadow.Value };

        if (Color is not null)
            nf = nf with { FontColor = Color.Key };

        if (Size is not null)
            nf = nf with { FontSize = Size.Value.Points };

        if (Underline is not null)
            nf = nf with { Underline = Underline.Value };

        if (VerticalAlignment is not null)
            nf = nf with { VerticalAlignment = VerticalAlignment.Value };

        if (Scheme is not null)
            nf = nf with { FontScheme = Scheme.Value };

        return nf;
    }
}
