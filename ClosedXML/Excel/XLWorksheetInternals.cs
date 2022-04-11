using System;

namespace ClosedXML.Excel
{
    internal class XLWorksheetInternals : IDisposable
    {
        private bool _disposed = false;

        public XLWorksheetInternals(
            XLCellsCollection cellsCollection,
            XLColumnsCollection columnsCollection,
            XLRowsCollection rowsCollection,
            XLRanges mergedRanges)
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
            // Dispose of unmanaged resources.
            Dispose(true);
            // Suppress finalization.
            GC.SuppressFinalize(this);
        }

        public void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
            {
                CellsCollection?.Clear();
                ColumnsCollection?.Clear();
                RowsCollection?.Clear();
                MergedRanges?.RemoveAll();
            }

            _disposed = true;
        }
    }
}
