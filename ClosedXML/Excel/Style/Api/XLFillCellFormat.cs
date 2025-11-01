using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// An API object to modify fill of a cell range.
/// </summary>
/// <remarks>
/// Most of the various logic that adjusts background color or pattern based on changes to pattern
/// type or background color have been inherited from original implementation.
/// </remarks>
internal sealed partial class XLFillCellFormat
{
    private static readonly XLPatternFill PatternNone = new()
    {
        PatternType = XLFillPatternValues.None,
        PatternColor = XLColor.Auto,
        BackgroundColor = XLColor.Auto
    };

    private readonly XLCellFormat _parent;

    public XLFillCellFormat(XLCellFormat parent)
    {
        _parent = parent;
    }

    private XLColor BackgroundColor
    {
        get => _parent.Resolve(static x => (x.Fill.Pattern ?? PatternNone).BackgroundColor);
        set => _parent.ModifyFill(static (fill, bgColor) =>
        {
            var currentFill = fill.Pattern ?? PatternNone;

            XLFillPatternValues newPattern;
            if (bgColor.IsAuto && currentFill.PatternType == XLFillPatternValues.Solid)
            {
                // Fill is single color, and that color will be is transparent -> there is no fill.
                newPattern = XLFillPatternValues.None;
            }
            else if (!bgColor.IsAuto && currentFill.PatternType == XLFillPatternValues.None)
            {
                // Color is not transparent, but fill is none -> make it solid
                newPattern = XLFillPatternValues.Solid;
            }
            else
            {
                newPattern = currentFill.PatternType;
            }

            return new XLFillFormatValue(new XLPatternFill
            {
                PatternType = newPattern,
                PatternColor = currentFill.PatternColor,
                BackgroundColor = bgColor
            });
        }, value);
    }

    private XLColor PatternColor
    {
        get => _parent.Resolve(x => (x.Fill.Pattern ?? PatternNone).PatternColor);
        set => _parent.ModifyFill((fill, patternColor) =>
        {
            var pattern = fill.Pattern ?? PatternNone;
            return new XLFillFormatValue(pattern with { PatternColor = patternColor });
        }, value);
    }

    private XLFillPatternValues PatternType
    {
        get => _parent.Resolve(static x => (x.Fill.Pattern ?? PatternNone).PatternType);
        set => _parent.ModifyFill(static (fill, patternType) =>
        {
            var pattern = fill.Pattern ?? PatternNone;
            if (pattern.PatternType == XLFillPatternValues.None && patternType != XLFillPatternValues.None)
            {
                // Keep the original behavior. Just having a non-none pattern type doesn't actually
                // affects color of cells visually, it needs at least background color.
                return new XLFillFormatValue(pattern with
                {
                    BackgroundColor = XLColor.FromTheme(XLThemeColor.Text1),
                    PatternType = patternType
                });
            }

            return new XLFillFormatValue(pattern with { PatternType = patternType });
        }, value);
    }

    internal void SetValue(IXLFill value)
    {
        // No need for shenanigans with changing pattern fill or colors, the original should be valid and consistent.
        _parent.ModifyFill(static (_, patternFill) => new XLFillFormatValue(new XLPatternFill
        {
            PatternType = patternFill.PatternType,
            PatternColor = patternFill.PatternColor,
            BackgroundColor = patternFill.BackgroundColor,
        }), value);
    }
}
