namespace ClosedXML.Excel.Formatting;

/// <summary>
/// A fill formatting. Fill can have at most one pattern that specifies how to fill a cell.
/// </summary>
internal record XLFillFormat
{
    public static readonly XLFillFormat Empty = new(null, null, null);

    public XLFillFormat(XLPatternFill pattern)
        : this(pattern, null, null)
    {
    }

    public XLFillFormat(XLLinearGradientFill linearGradient)
        : this(null, linearGradient, null)
    {
    }

    public XLFillFormat(XLPathGradientFill pathGradient)
        : this(null, null, pathGradient)
    {
    }

    private XLFillFormat(XLPatternFill? pattern, XLLinearGradientFill? linearGradient, XLPathGradientFill? pathGradient)
    {
        Pattern = pattern;
        LinearGradient = linearGradient;
        PathGradient = pathGradient;
    }

    public XLPatternFill? Pattern { get; }

    public XLLinearGradientFill? LinearGradient { get; }

    public XLPathGradientFill? PathGradient { get; }

    internal XLFillKey ApplyTo(XLFillKey fillKey)
    {
        // TODO: XLFillKey doesn't have structure to hold other gradient types. For now, keep original logic.
        if (Pattern is null)
            return fillKey;

        switch (Pattern.PatternType)
        {
            case XLFillPatternValues.Solid:
                // ISO-29500: For solid cell fills (no pattern), fgColor is used.
                // That makes sense, because solid pattern means the pattern color fully covers
                // everything, but the historical ClosedXML code expects color for solid to be
                // in the background color, so keep it for now.
                fillKey = new XLFillKey
                {
                    PatternType = Pattern.PatternType,
                    PatternColor = XLColor.NoColor.Key,
                    BackgroundColor = Pattern.PatternColor.Key
                };
                break;
            default:
                fillKey = new XLFillKey
                {
                    PatternType = Pattern.PatternType,
                    PatternColor = Pattern.PatternColor.Key,
                    BackgroundColor = Pattern.BackgroundColor.Key
                };
                break;
        }

        return fillKey;
    }
}

