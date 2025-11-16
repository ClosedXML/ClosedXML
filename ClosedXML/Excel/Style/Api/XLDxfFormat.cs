using System;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

internal partial class XLDxFormat
{
    private readonly XLWorkbookStyles _styles;
    private readonly IXLDxfContainer _container;

    public XLDxFormat(XLWorkbookStyles styles, IXLDxfContainer container)
    {
        _styles = styles;
        _container = container;
    }

    private XLDxfValue Dxf => _container.FormatValue ?? XLDxfValue.Empty;

    private XLDxfFontFormat Font => new(this);

    internal TProperty? Resolve<TComponent, TProperty>(Func<XLDxfValue, TComponent> getComponent, Func<TComponent, TProperty?> getProperty)
        where TProperty : struct
    {
        var component = getComponent(Dxf);
        return getProperty(component);
    }

    internal TProperty? Resolve<TComponent, TProperty>(Func<XLDxfValue, TComponent> getComponent, Func<TComponent, TProperty?> getProperty)
        where TProperty : class
    {
        var component = getComponent(Dxf);
        return getProperty(component);
    }

    internal void ModifyFont<T>(Func<XLDifferentialFontValue, T, XLDifferentialFontValue> modify, T value)
    {
        var modifiedDxf = _styles.GetRegisteredDxFormat(Dxf, dxf =>
        {
            var modifiedFont = modify(dxf.Font, value);
            var modifiedDxf = dxf with { Font = modifiedFont };
            return modifiedDxf;
        });
        _container.FormatValue = modifiedDxf;
    }

    internal void SetValue(IXLStyle value)
    {
        throw new NotImplementedException();
    }
}
