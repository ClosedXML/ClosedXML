using System;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// An API object to modify a font format.
/// </summary>
internal class XLFontFormat : IXLFontFormat<IXLDifferentialFormat>
{
    private readonly IXLDifferentialFormat _parent;
    private readonly IDxfContainer _container;
    private readonly XLWorkbookStyles _styles;

    internal XLFontFormat(IXLDifferentialFormat parent, IDxfContainer container, XLWorkbookStyles styles)
    {
        _parent = parent;
        _container = container;
        _styles = styles;
    }

    public bool? Bold
    {
        get => FontFormat?.Bold;
        set => SetFont(f => f with { Bold = value });
    }

    internal XLFontFormatValue? FontFormat => _container.FormatValue.Font;

    public IXLDifferentialFormat SetBold()
    {
        return SetBold(true);
    }

    public IXLDifferentialFormat SetBold(bool? value)
    {
        Bold = value;
        return _parent;
    }

    private void SetFont(Func<XLFontFormatValue, XLFontFormatValue> modifyFont)
    {
        var fontFormat = _styles.GetRegisteredFontFormat(FontFormat ?? XLFontFormatValue.Empty, modifyFont);
        var dxf = _styles.GetRegisteredDxFormat(_container.FormatValue, format => format with { Font = fontFormat });
        _container.FormatValue = dxf;
    }
}
