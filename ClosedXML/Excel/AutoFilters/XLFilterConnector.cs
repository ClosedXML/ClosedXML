namespace ClosedXML.Excel;

internal class XLFilterConnector : IXLFilterConnector
{
    private readonly XLAutoFilter _autoFilter;
    private readonly XLFilterColumn _filterColumn;

    public XLFilterConnector(XLAutoFilter autoFilter, XLFilterColumn filterColumn)
    {
        _autoFilter = autoFilter;
        _filterColumn = filterColumn;
    }

    public IXLCustomFilteredColumn And => new XLCustomFilteredColumn(_autoFilter, _filterColumn, XLConnector.And);

    public IXLCustomFilteredColumn Or => new XLCustomFilteredColumn(_autoFilter, _filterColumn, XLConnector.Or);
}
