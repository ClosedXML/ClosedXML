using System;

namespace ClosedXML.Excel;

internal sealed class XLWorksheetInternals : IDisposable
{
    private bool _disposed;

    internal required XLCellsCollection CellsCollection
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        init;
    }

    internal required XLColumnsCollection ColumnsCollection
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        init;
    }

    internal required XLRowsCollection RowsCollection
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        init;
    }

    internal required XLRanges MergedRanges
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        init;
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        CellsCollection.ValueSlice.DereferenceSlice();
        CellsCollection.Clear();
        ColumnsCollection.Clear();
        RowsCollection.Clear();
        MergedRanges.RemoveAll();
        _disposed = true;
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(XLWorksheetInternals));
    }
}
