
namespace ClosedXML.Excel
{
    internal interface IXLWorksheetInternals
    {
        XLCellCollection CellsCollection { get; }
        XLColumnsCollection ColumnsCollection { get; }
        XLRowsCollection RowsCollection { get; }
        XLRanges MergedRanges { get; }
        XLWorkbook Workbook { get; }
    }
}
