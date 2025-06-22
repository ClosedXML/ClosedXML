using System;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// API object to modify font properties of a cell format of a <see cref="IXLFormatContainer"/>.
/// </summary>
internal class XLFontCellFormat
{
    private readonly XLCellFormat _parent;
    private readonly FormatHierarchy _hierarchy;
    private readonly XLWorkbookStyles _styles;

    internal XLFontCellFormat(XLCellFormat parent, FormatHierarchy hierarchy, XLWorkbookStyles styles)
    {
        _parent = parent;
        _hierarchy = hierarchy;
        _styles = styles;
    }

    public XLFontName Name => Resolve(static x => x.Font?.Name);

    public bool Bold => Resolve(static x => x.Font?.Bold);

    public bool Italic => Resolve(static x => x.Font?.Italic);

    /// <summary>
    /// Size in points.
    /// </summary>
    public double Size => Resolve(static x => x.Font?.Size).Points;

    private T Resolve<T>(Func<XLCellFormatValue, T?> selector)
        where T : struct
    {
        return _hierarchy.Resolve(selector, _styles.DefaultFormat);
    }
}
