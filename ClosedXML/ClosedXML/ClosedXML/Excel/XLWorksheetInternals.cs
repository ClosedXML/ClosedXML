namespace ClosedXML.Excel
{
    internal class XLWorksheetInternals
    {
        public XLWorksheetInternals(
            XLCellsCollection cellsCollection, 
            XLColumnsCollection columnsCollection,
            XLRowsCollection rowsCollection,
            XLRanges mergedRanges,
            XLWorkbook workbook
            )
        {
            CellsCollection = cellsCollection;
            ColumnsCollection = columnsCollection;
            RowsCollection = rowsCollection;
            MergedRanges = mergedRanges;
            Workbook = workbook;
        }

        public XLCellsCollection CellsCollection { get; private set; }
        public XLColumnsCollection ColumnsCollection { get; private set; }
        public XLRowsCollection RowsCollection { get; private set; }
        public XLRanges MergedRanges { get; internal set; }
        public XLWorkbook Workbook { get; internal set; }
    }
}
