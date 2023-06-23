#nullable disable

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

        public XLCellsCollection CellsCollection { get; }
        public XLColumnsCollection ColumnsCollection { get; }
        public XLRowsCollection RowsCollection { get; }
        public XLRanges MergedRanges { get; internal set; }

        // Used by Janitor.Fody
        private void DisposeManaged()
        {
            CellsCollection.ValueSlice.DereferenceSlice();
            CellsCollection.Clear();
            ColumnsCollection.Clear();
            RowsCollection.Clear();
            MergedRanges.RemoveAll();
        }

        public void Dispose()
        {
            // Leave this empty so that Janitor.Fody can do its work
        }
    }
}
