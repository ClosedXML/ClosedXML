namespace ClosedXML.Excel;

internal class XLFilterConnector : IXLFilterConnector
{
    private readonly XLFilterColumn _filterColumn;

    public XLFilterConnector(XLFilterColumn filterColumn)
    {
        _filterColumn = filterColumn;
    }

    public IXLCustomFilteredColumn And => new XLCustomFilteredColumn(_filterColumn, XLConnector.And);

    public IXLCustomFilteredColumn Or => new XLCustomFilteredColumn(_filterColumn, XLConnector.Or);
}
