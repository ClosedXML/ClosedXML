using System;

namespace ClosedXML.Excel;

internal class XLPivotStyleFormat : XLPivotStyleFormatBase
{
    private readonly Func<XLPivotArea, bool> _filter;
    private readonly Func<XLPivotArea> _factory;

    public XLPivotStyleFormat(XLPivotTable pivotTable, Func<XLPivotArea, bool> filter, Func<XLPivotArea> factory)
        : base(pivotTable)
    {
        _filter = filter;
        _factory = factory;
    }

    internal override XLPivotArea GetCurrentArea()
    {
        return _factory();
    }

    internal override bool Filter(XLPivotArea area)
    {
        return _filter(area);
    }
}
