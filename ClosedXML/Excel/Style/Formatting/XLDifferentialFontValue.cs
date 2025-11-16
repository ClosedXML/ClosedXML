namespace ClosedXML.Excel.Formatting;

/// <summary>
/// A differential font format.
/// </summary>
internal record XLDifferentialFontValue
{
    /// <summary>
    /// A value with all properties null.
    /// </summary>
    internal static XLDifferentialFontValue Empty = new()
    {
        Name = null,
        Charset = null,
        Family = null,
        Bold = null,
        Italic = null,
        Strikethrough = null,
        Outline = null,
        Shadow = null,
        Condense = null,
        Extend = null,
        Color = null,
        Size = null,
        Underline = null,
        VerticalAlignment = null,
        Scheme = null,
    };

    internal required XLFontName? Name { get; init; }

    internal required XLFontCharSet? Charset { get; init; }

    internal required XLFontFamilyNumberingValues? Family { get; init; }

    internal required bool? Bold { get; init; }

    internal required bool? Italic { get; init; }

    internal required bool? Strikethrough { get; init; }

    internal required bool? Outline { get; init; }

    internal required bool? Shadow { get; init; }

    internal required bool? Condense { get; init; }

    internal required bool? Extend { get; init; }

    internal required XLColor? Color { get; init; }

    internal required XLFontSize? Size { get; init; }

    internal required XLFontUnderlineValues? Underline { get; init; }

    internal required XLFontVerticalTextAlignmentValues? VerticalAlignment { get; init; }

    internal required XLFontScheme? Scheme { get; init; }

    internal XLFontKey ApplyTo(XLFontKey key)
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

    internal bool IsEmpty()
    {
        return Name is null &&
               Charset is null &&
               Family is null &&
               Bold is null &&
               Italic is null &&
               Strikethrough is null &&
               Outline is null &&
               Shadow is null &&
               Condense is null &&
               Extend is null &&
               Color is null &&
               Size is null &&
               Underline is null &&
               VerticalAlignment is null &&
               Scheme is null;
    }
}
