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

        // Used by Janitor.Fody
        private void DisposeManaged()
        {
            CellsCollection.Clear();
            ColumnsCollection.Clear();
            RowsCollection.Clear();
            MergedRanges.RemoveAll();
        }

#if _NET40_

        public void Dispose()
        {
            // net40 doesn't support Janitor.Fody, so let's dispose manually
            DisposeManaged();
        }

#else

        public void Dispose()
        {
            // Leave this empty (for non net40 targets) so that Janitor.Fody can do its work
        }

#endif
    }
}
