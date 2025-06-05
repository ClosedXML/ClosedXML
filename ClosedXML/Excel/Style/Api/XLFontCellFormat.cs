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

    public XLFontName Name => _hierarchy.Resolve(static x => x.Font?.Name, _styles.DefaultFormat);

    public bool Bold => _hierarchy.Resolve(static x => x.Font?.Bold, _styles.DefaultFormat);

    public bool Italic => _hierarchy.Resolve(static x => x.Font?.Italic, _styles.DefaultFormat);

    /// <summary>
    /// Size in points.
    /// </summary>
    public double Size => _hierarchy.Resolve(static x => x.Font?.Size, _styles.DefaultFormat).Points;
}
