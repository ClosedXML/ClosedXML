using System;

namespace ClosedXML.Excel;

/// <summary>
/// API object to modify font properties of a cell format of a <see cref="IXLFormatContainer"/>.
/// Unlike the <see cref="XLStyle"/>, the <see cref="XLCellFormat"/> one modifies formatting
/// in a <see cref="XLWorkbookStyles"/>.
/// </summary>
internal class XLCellFormat
{
    private readonly FormatHierarchy _hierarchy;
    private readonly XLWorkbookStyles _styles;

    internal XLCellFormat(FormatHierarchy hierarchy, XLWorkbookStyles styles)
    {
        _hierarchy = hierarchy;
        _styles = styles ?? throw new ArgumentNullException(nameof(styles));
    }

    internal XLFontCellFormat Font => new(this, _hierarchy, _styles);
}
