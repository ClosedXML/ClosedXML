
namespace ClosedXML.Excel
{
    internal interface IXLWorksheetInternals
    {
        XLCellCollection CellsCollection { get; }
        XLColumnsCollection ColumnsCollection { get; }
        XLRowsCollection RowsCollection { get; }
        XLMergedRanges MergedRanges { get; }
        XLWorkbook Workbook { get; }
    }
}
