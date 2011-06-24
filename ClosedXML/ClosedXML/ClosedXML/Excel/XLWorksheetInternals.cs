namespace ClosedXML.Excel
{
    internal class XLWorksheetInternals: IXLWorksheetInternals
    {
        public XLWorksheetInternals(
            XLCellCollection cellsCollection, 
            XLColumnsCollection columnsCollection,
            XLRowsCollection rowsCollection,
            XLMergedRanges mergedRanges,
            XLWorkbook workbook
            )
        {
            CellsCollection = cellsCollection;
            ColumnsCollection = columnsCollection;
            RowsCollection = rowsCollection;
            MergedRanges = mergedRanges;
            Workbook = workbook;
        }

        public XLCellCollection CellsCollection { get; private set; }
        public XLColumnsCollection ColumnsCollection { get; private set; }
        public XLRowsCollection RowsCollection { get; private set; }
        public XLMergedRanges MergedRanges { get; internal set; }
        public XLWorkbook Workbook { get; internal set; }
    }
}
