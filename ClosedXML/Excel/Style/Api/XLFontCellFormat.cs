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

    public bool Bold => _hierarchy.Resolve(static x => x.Font?.Bold, false);

    public bool Italic => _hierarchy.Resolve(static x => x.Font?.Italic, false);

    public XLFontName Name => _hierarchy.ResolveWithNormalFallback(static x => x.Font?.Name);
}
