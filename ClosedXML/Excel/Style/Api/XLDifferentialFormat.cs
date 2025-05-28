using System;

namespace ClosedXML.Excel;

/// <summary>
/// API object to modify differential format of a <see cref="IDxfContainer"/>.
/// </summary>
internal class XLDifferentialFormat : IXLDifferentialFormat
{
    private readonly IDxfContainer _container;
    private readonly XLWorkbookStyles _styles;

    public XLDifferentialFormat(IDxfContainer container, XLWorkbookStyles styles)
    {
        _container = container ?? throw new ArgumentNullException(nameof(container));
        _styles = styles ?? throw new ArgumentNullException(nameof(styles));
    }

    public IXLFontFormat<IXLDifferentialFormat> Font => new XLFontFormat(this, _container, _styles);
}
