using System;

namespace ClosedXML.Excel
{
    internal class XLWorksheetInternals : IDisposable
    {
        public XLWorksheetInternals(
            XLCellsCollection cellsCollection,
            XLColumnsCollection columnsCollection,
            XLRowsCollection rowsCollection,
            XLRanges mergedRanges
            )
        {
            CellsCollection = cellsCollection;
            ColumnsCollection = columnsCollection;
            RowsCollection = rowsCollection;
            MergedRanges = mergedRanges;
        }

        public XLCellsCollection CellsCollection { get; private set; }
        public XLColumnsCollection ColumnsCollection { get; private set; }
        public XLRowsCollection RowsCollection { get; private set; }
        public XLRanges MergedRanges { get; internal set; }

        public void Dispose()
        {
            ColumnsCollection.Dispose();
            RowsCollection.Dispose();
            MergedRanges.Dispose();
        }
    }
}
