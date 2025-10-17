namespace ClosedXML.Excel.Formatting;

/// <summary>
/// A formatting record for <see cref="XLCellFormatValue"/>.
/// </summary>
internal record XLFontFormatValue
{
    /// <summary>
    /// <para>
    /// A default font values.
    /// <list type="bullet">
    ///   <item>
    ///   When a font is loaded from the styles part, unspecified props use these values. The values
    ///   are essentially reinterpreted <em>zero</em> (other than name and size that can't be zero).
    ///   </item>
    ///   <item>
    ///   When font property has this value, it is omitted from being saved to font element (other
    ///   than font and size). Font element is pretty buggy, but these are roughly XSD default values.
    ///   </item>
    /// </list>
    /// The default values are different from normal style. Normal style is what should default
    /// format look like for user when a new workbook is created. The default font is about loading and
    /// saving.
    /// </para>
    /// <para>
    /// The font name and size are special, but the values in the property are the default fallback
    /// when even default format doesn't contain font name and size.
    /// </para>
    /// </summary>
    public static readonly XLFontFormatValue Default = new()
    {
        Name = "Calibri",
        Charset = XLFontCharSet.Ansi,
        Family = XLFontFamilyNumberingValues.NotApplicable,
        Bold = false,
        Italic = false,
        Strikethrough = false,
        Outline = false,
        Shadow = false,
        Condense = false,
        Extend = false,
        Color = XLColor.Auto,
        Size = XLFontSize.FromPoints(11),
        Underline = XLFontUnderlineValues.None,
        VerticalAlignment = XLFontVerticalTextAlignmentValues.Baseline,
        Scheme = XLFontScheme.None
    };

    public required XLFontName Name { get; init; }

    public required XLFontCharSet Charset { get; init; }

    public required XLFontFamilyNumberingValues Family { get; init; }

    public required bool Bold { get; init; }

    public required bool Italic { get; init; }

    public required bool Strikethrough { get; init; }

    public required bool Outline { get; init; }

    public required bool Shadow { get; init; }

    public required bool Condense { get; init; }

    public required bool Extend { get; init; }

    public required XLColor Color { get; init; }

    public required XLFontSize Size { get; init; }

    public required XLFontUnderlineValues Underline { get; init; }

    public required XLFontVerticalTextAlignmentValues VerticalAlignment { get; init; }

    public required XLFontScheme Scheme { get; init; }

    internal XLFontKey GetFontKey()
    {
        // XLFontKey doesn't contain outline, condense or extend
        return new XLFontKey
        {
            FontName = Name.Text,
            FontCharSet = Charset,
            FontFamilyNumbering = Family,
            Bold = Bold,
            Italic = Italic,
            Strikethrough = Strikethrough,
            Shadow = Shadow,

            // TODO Styles: Incorrect default value for old XLFontValue.Color
            // The correct color is auto, but XLFontValue uses black (0xFF000000).
            // Changes few test workbooks, don't fix now. Fix it in bulk, when
            // switching to new styles infra.
            FontColor = Color == XLColor.Auto ? XLColor.FromArgb(0, 0, 0).Key : Color.Key,
            FontSize = Size.Points,
            Underline = Underline,
            VerticalAlignment = VerticalAlignment,
            FontScheme = Scheme
        };
    }
}
