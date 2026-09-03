using System;

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
        Color = XLColor.Automatic,
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

    /// <summary>
    /// Return a registered font format equivalent to <paramref name="font"/>.
    /// </summary>
    internal static XLFontFormatValue FromFontBase(IXLFontBase font, XLWorkbookStyles styles)
    {
        var fontFormat = new XLFontFormatValue
        {
            Name = font.FontName,
            Charset = font.FontCharSet,
            Family = font.FontFamilyNumbering,
            Bold = font.Bold,
            Italic = font.Italic,
            Strikethrough = font.Strikethrough,
            Outline = Default.Outline,
            Shadow = font.Shadow,
            Condense = Default.Condense,
            Extend = Default.Extend,
            Color = font.FontColor,
            Size = XLFontSize.FromPoints(font.FontSize),
            Underline = font.Underline,
            VerticalAlignment = font.VerticalAlignment,
            Scheme = font.FontScheme
        };
        return styles.GetRegisteredFontFormat(fontFormat, static x => x);
    }

    /// <summary>
    /// Create an adapter to font base. The adapter is not modifiable.
    /// </summary>
    internal IXLFontBase ToFontBase()
    {
        return new FontBaseAdapter(this);
    }

    /// <summary>
    /// Return a font that will have its values taken from the dxf where the dxf does have value.
    /// </summary>
    internal XLFontFormatValue AdjustWith(XLDifferentialFontValue dxf)
    {
        if (dxf.IsEmpty())
            return this;

        var result = this;
        if (dxf.Name is { } name)
            result = result with { Name = name };

        if (dxf.Charset is { } charset)
            result = result with { Charset = charset };

        if (dxf.Family is { } family)
            result = result with { Family = family };

        if (dxf.Bold is { } bold)
            result = result with { Bold = bold };

        if (dxf.Italic is { } italic)
            result = result with { Italic = italic };

        if (dxf.Strikethrough is { } strikethrough)
            result = result with { Strikethrough = strikethrough };

        if (dxf.Outline is { } outline)
            result = result with { Outline = outline };

        if (dxf.Shadow is { } shadow)
            result = result with { Shadow = shadow };

        if (dxf.Condense is { } condense)
            result = result with { Condense = condense };

        if (dxf.Extend is { } extend)
            result = result with { Extend = extend };

        if (dxf.Color is { } color)
            result = result with { Color = color };

        if (dxf.Size is { } size)
            result = result with { Size = size };

        if (dxf.Underline is { } underline)
            result = result with { Underline = underline };

        if (dxf.VerticalAlignment is { } verticalAlignment)
            result = result with { VerticalAlignment = verticalAlignment };

        if (dxf.Scheme is { } scheme)
            result = result with { Scheme = scheme };

        return result;
    }

    private class FontBaseAdapter : IXLFontBase
    {
        private readonly XLFontFormatValue _font;

        internal FontBaseAdapter(XLFontFormatValue font)
        {
            _font = font;
        }

        public bool Bold
        {
            get => _font.Bold;
            set => throw Exception();
        }

        public bool Italic
        {
            get => _font.Italic;
            set => throw Exception();
        }
        public XLFontUnderlineValues Underline
        {
            get => _font.Underline;
            set => throw Exception();
        }

        public bool Strikethrough
        {
            get => _font.Strikethrough;
            set => throw Exception();
        }

        public XLFontVerticalTextAlignmentValues VerticalAlignment
        {
            get => _font.VerticalAlignment;
            set => throw Exception();
        }

        public bool Shadow
        {
            get => _font.Shadow;
            set => throw Exception();
        }

        public double FontSize
        {
            get => _font.Size.Points;
            set => throw Exception();
        }

        public XLColor FontColor
        {
            get => _font.Color;
            set => throw Exception();
        }

        public string FontName
        {
            get => _font.Name.Text;
            set => throw Exception();
        }

        public XLFontFamilyNumberingValues FontFamilyNumbering
        {
            get => _font.Family;
            set => throw Exception();
        }

        public XLFontCharSet FontCharSet
        {
            get => _font.Charset;
            set => throw Exception();
        }

        public XLFontScheme FontScheme
        {
            get => _font.Scheme;
            set => throw Exception();
        }

        private NotSupportedException Exception()
        {
            return new NotSupportedException("This is an adapter for an immutable font format.");
        }
    }
}
