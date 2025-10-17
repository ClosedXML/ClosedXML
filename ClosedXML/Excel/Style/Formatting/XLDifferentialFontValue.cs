namespace ClosedXML.Excel.Formatting;

/// <summary>
/// A differential font format.
/// </summary>
internal record XLDifferentialFontValue
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

    public XLFontKey ApplyTo(XLFontKey key)
    {
        // XLFontKey doesn't contain outline, condense or extend
        if (Name is not null)
            key = key with { FontName = Name.Value.Text };

        if (Charset is not null)
            key = key with { FontCharSet = Charset.Value };

        if (Family is not null)
            key = key with { FontFamilyNumbering = Family.Value };

        if (Bold is not null)
            key = key with { Bold = Bold.Value };

        if (Italic is not null)
            key = key with { Italic = Italic.Value };

        if (Strikethrough is not null)
            key = key with { Strikethrough = Strikethrough.Value };

        if (Shadow is not null)
            key = key with { Shadow = Shadow.Value };

        if (Color is not null)
            key = key with { FontColor = Color.Key };

        if (Size is not null)
            key = key with { FontSize = Size.Value.Points };

        if (Underline is not null)
            key = key with { Underline = Underline.Value };

        if (VerticalAlignment is not null)
            key = key with { VerticalAlignment = VerticalAlignment.Value };

        if (Scheme is not null)
            key = key with { FontScheme = Scheme.Value };

        return key;
    }
}
