// Keep this file CodeMaid organised and cleaned
namespace ClosedXML.Excel;

internal class XLPivotTableStyleFormats : IXLPivotTableStyleFormats
{
    private readonly XLPivotTable _pivotTable;

    public XLPivotTableStyleFormats(XLPivotTable pivotTable)
    {
        _pivotTable = pivotTable;
    }

    #region IXLPivotTableStyleFormats members

    public IXLPivotStyleFormats ColumnGrandTotalFormats => new XLPivotStyleFormats(_pivotTable, isRowGrand: false);

    public IXLPivotStyleFormats RowGrandTotalFormats => new XLPivotStyleFormats(_pivotTable, isRowGrand: true);

    #endregion IXLPivotTableStyleFormats members
}
