#nullable disable

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Slice is a sparse array that stores a part of cell information (e.g. only values,
    /// only styles ...). Slice has same size as a worksheet. If some cells are pushed out
    /// of the permitted range, they are gone.
    /// </summary>
    /// <remarks>
    /// This is a ref return, so if the underlaying value
    /// changes, the returned value also changes. To avoid,
    /// just don't use <c>ref</c> and structs will be copied.
    /// </remarks>
    /// <typeparam name="TElement">The type of data stored in the slice.</typeparam>
    internal partial class Slice<TElement> : ISlice
    {
        private static readonly Lut<TElement> Dummy = new();
        private readonly TElement _defaultValue = default;

        /// <summary>
        /// The content of the slice. Note that LUT uses index that starts from 0,
        /// so rows and columns must be adjusted to retrieved the value.
        /// </summary>
        private readonly Lut<Lut<TElement>> _data;

        /// <summary>
        /// Key is column number, value is number of cells in the column that are used.
        /// </summary>
        private readonly Dictionary<int, int> _columnUsage = new();

        internal Slice()
        {
            _data = new();
        }

        /// <summary>
        /// Get the slice value at the specified point of the sheet.
        /// </summary>
        internal ref readonly TElement this[XLSheetPoint point] => ref this[point.Row, point.Column];

        /// <summary>
        /// Get the slice value at the specified point of the sheet.
        /// </summary>
        internal ref readonly TElement this[int row, int column]
        {
            get
            {
                var rowLut = _data.Get(row - 1);
                if (rowLut is null)
                    return ref _defaultValue;

                return ref rowLut.Get(column - 1);
            }
        }

        /// <inheritdoc />
        public bool IsEmpty => MaxRow == 0;

        /// <inheritdoc />
        public int MaxColumn { get; private set; }

        /// <inheritdoc />
        public int MaxRow => _data.MaxUsedIndex + 1;

        /// <inheritdoc />
        public IEnumerable<int> UsedRows
        {
            get
            {
                var rowsEnumerator = new Lut<Lut<TElement>>.LutEnumerator(_data, XLHelper.MinRowNumber - 1, XLHelper.MaxRowNumber - 1);
                while (rowsEnumerator.MoveNext())
                {
                    if (!rowsEnumerator.Current.IsEmpty)
                        yield return rowsEnumerator.Index + 1;
                }
            }
        }

        /// <inheritdoc />
        public Dictionary<int, int>.KeyCollection UsedColumns => _columnUsage.Keys;

        /// <inheritdoc />
        public void Clear(XLSheetRange range)
        {
            var enumerator = new Enumerator(this, range);
            while (enumerator.MoveNext())
            {
                Set(enumerator.Point, in _defaultValue);
            }
        }

        /// <inheritdoc />
        public void DeleteAreaAndShiftLeft(XLSheetRange rangeToDelete)
        {
            Clear(rangeToDelete);

            var noCellsToShift = rangeToDelete.LastPoint.Column == XLHelper.MaxColumnNumber;
            if (noCellsToShift)
                return;

            var shiftDistance = rangeToDelete.Width;
            var shiftRange = rangeToDelete.RightRange();
            var cellEnumerator = new Enumerator(this, shiftRange);
            while (cellEnumerator.MoveNext())
            {
                var srcPoint = cellEnumerator.Point;
                var dstPoint = new XLSheetPoint(srcPoint.Row, srcPoint.Column - shiftDistance);
                Set(dstPoint, in cellEnumerator.Current);
                Set(srcPoint, in _defaultValue);
            }
        }

        /// <inheritdoc />
        public void DeleteAreaAndShiftUp(XLSheetRange rangeToDelete)
        {
            Clear(rangeToDelete);

            var noCellsToShift = rangeToDelete.LastPoint.Row == XLHelper.MaxRowNumber;
            if (noCellsToShift)
                return;

            var shiftDistance = rangeToDelete.Height;
            var shiftRange = rangeToDelete.BelowRange();
            var cellEnumerator = new Enumerator(this, shiftRange);
            while (cellEnumerator.MoveNext())
            {
                var srcPoint = cellEnumerator.Point;
                var dstPoint = new XLSheetPoint(srcPoint.Row - shiftDistance, srcPoint.Column);
                Set(dstPoint, in cellEnumerator.Current);
                Set(srcPoint, in _defaultValue);
            }
        }

        /// <summary>
        /// Get enumerator over used values of the range.
        /// </summary>
        public IEnumerator<XLSheetPoint> GetEnumerator(XLSheetRange range, bool reverse = false)
        {
            return !reverse ? new Enumerator(this, range) : new ReverseEnumerator(this, range);
        }

        /// <inheritdoc />
        public void InsertAreaAndShiftDown(XLSheetRange range)
        {
            var hasSpaceBelow = range.LastPoint.Row < XLHelper.MaxRowNumber;
            if (!hasSpaceBelow)
            {
                Clear(range);
                return;
            }

            var shiftDistance = range.Height;

            // Purged range might contain some cells that wouldn't be overwritten during shift => clear.
            var purgedRange = new XLSheetRange(
                new XLSheetPoint(XLHelper.MaxRowNumber - shiftDistance + 1, range.FirstPoint.Column),
                new XLSheetPoint(XLHelper.MaxRowNumber, range.LastPoint.Column));
            Clear(purgedRange);

            var shiftedRange = new XLSheetRange(
                range.FirstPoint,
                new XLSheetPoint(XLHelper.MaxRowNumber - shiftDistance, range.LastPoint.Column));
            var cellEnumerator = new ReverseEnumerator(this, shiftedRange);
            while (cellEnumerator.MoveNext())
            {
                var srcPoint = cellEnumerator.Point;
                var dstPoint = new XLSheetPoint(srcPoint.Row + shiftDistance, srcPoint.Column);
                Set(dstPoint, in cellEnumerator.Current);
                Set(srcPoint, in _defaultValue);
            }
        }

        /// <inheritdoc />
        public void InsertAreaAndShiftRight(XLSheetRange range)
        {
            var hasSpaceRight = range.LastPoint.Column < XLHelper.MaxColumnNumber;
            if (!hasSpaceRight)
            {
                Clear(range);
                return;
            }

            var shiftDistance = range.Width;

            // Purged range might contain some cells that wouldn't be overwritten during shift => clear.
            var purgedRange = new XLSheetRange(
                new XLSheetPoint(range.FirstPoint.Row, XLHelper.MaxColumnNumber - shiftDistance + 1),
                new XLSheetPoint(range.LastPoint.Row, XLHelper.MaxColumnNumber));
            Clear(purgedRange);

            var shiftedRange = new XLSheetRange(
                range.FirstPoint,
                new XLSheetPoint(range.LastPoint.Row, XLHelper.MaxColumnNumber - shiftDistance));
            var enumerator = new ReverseEnumerator(this, shiftedRange);
            while (enumerator.MoveNext())
            {
                var srcPoint = enumerator.Point;
                var dstPoint = new XLSheetPoint(srcPoint.Row, srcPoint.Column + shiftDistance);
                Set(dstPoint, in enumerator.Current);
                Set(srcPoint, in _defaultValue);
            }
        }

        public bool IsUsed(XLSheetPoint address)
        {
            var rowLut = _data.Get(address.Row - 1);
            if (rowLut is null)
                return false;

            return rowLut.IsUsed(address.Column - 1);
        }

        public void Swap(XLSheetPoint sp1, XLSheetPoint sp2)
        {
            var value1 = this[sp1];
            var value2 = this[sp2];
            Set(sp1, in value2);
            Set(sp2, in value1);
        }

        internal void Set(XLSheetPoint point, in TElement value)
            => Set(point.Row, point.Column, in value);

        internal void Set(int row, int column, in TElement value)
        {
            var rowLut = _data.Get(row - 1);
            if (rowLut is null)
            {
                rowLut = new Lut<TElement>();
                _data.Set(row - 1, rowLut);
            }

            var wasUsed = rowLut.IsUsed(column - 1);
            rowLut.Set(column - 1, value);
            var isUsed = rowLut.IsUsed(column - 1);

            if (wasUsed && !isUsed)
            {
                var newCount = DecrementColumnUsage(column);
                if (newCount == 0 && MaxColumn == column)
                {
                    MaxColumn = CalculateMaxColumn();
                }

                if (rowLut.IsEmpty)
                    _data.Set(row - 1, null);
            }

            if (!wasUsed && isUsed)
            {
                IncrementColumnUsage(column);
                if (column > MaxColumn)
                    MaxColumn = column;
            }
        }

        private int CalculateMaxColumn()
        {
            var maxColIdx = -1;
            var rowEnumerator = new Lut<Lut<TElement>>.LutEnumerator(_data, XLHelper.MinRowNumber - 1, XLHelper.MaxRowNumber - 1);
            while (rowEnumerator.MoveNext())
                maxColIdx = Math.Max(maxColIdx, rowEnumerator.Current.MaxUsedIndex);

            return maxColIdx + 1;
        }

        private int DecrementColumnUsage(int column)
        {
            if (!_columnUsage.TryGetValue(column, out var count))
                return 0;

            if (count > 1)
                return _columnUsage[column] = count - 1;

            _columnUsage.Remove(column);
            return 0;
        }

        private void IncrementColumnUsage(int column)
        {
            if (_columnUsage.TryGetValue(column, out var value))
                _columnUsage[column] = value + 1;
            else
                _columnUsage.Add(column, 1);
        }

        /// <summary>
        /// Enumerator that returns used values from a specified range.
        /// </summary>
        [DebuggerDisplay("{Point}:{Current}")]
        internal class Enumerator : IEnumerator<XLSheetPoint>
        {
            private readonly XLSheetRange _range;
            private Lut<TElement>.LutEnumerator _columnsEnumerator;
            private Lut<Lut<TElement>>.LutEnumerator _rowsEnumerator;

            internal Enumerator(Slice<TElement> slice, XLSheetRange range)
            {
                _range = range;

                _columnsEnumerator = new Lut<TElement>.LutEnumerator(Dummy, XLHelper.MaxColumnNumber + 1, XLHelper.MaxColumnNumber + 1);
                _rowsEnumerator = new Lut<Lut<TElement>>.LutEnumerator(
                    slice._data,
                    range.FirstPoint.Row - 1,
                    range.LastPoint.Row - 1);
            }

            public ref readonly TElement Current => ref _columnsEnumerator.Current;

            public XLSheetPoint Point => new(_rowsEnumerator.Index + 1, _columnsEnumerator.Index + 1);

            /// <summary>
            /// The movement is columns first, then rows.
            /// </summary>
            public bool MoveNext()
            {
                while (!_columnsEnumerator.MoveNext())
                {
                    if (!_rowsEnumerator.MoveNext())
                        return false;

                    _columnsEnumerator = new Lut<TElement>.LutEnumerator(
                        _rowsEnumerator.Current,
                        _range.FirstPoint.Column - 1,
                        _range.LastPoint.Column - 1);
                }

                return true;
            }

            void IEnumerator.Reset() => throw new NotSupportedException();

            XLSheetPoint IEnumerator<XLSheetPoint>.Current => Point;

            object IEnumerator.Current => Point;

            void IDisposable.Dispose() { }
        }

        [DebuggerDisplay("{Point}:{Current}")]
        private class ReverseEnumerator : IEnumerator<XLSheetPoint>
        {
            private readonly XLSheetRange _range;
            private Lut<TElement>.ReverseLutEnumerator _columnsEnumerator;
            private Lut<Lut<TElement>>.ReverseLutEnumerator _rowsEnumerator;

            internal ReverseEnumerator(Slice<TElement> slice, XLSheetRange range)
            {
                _range = range;
                _columnsEnumerator = new Lut<TElement>.ReverseLutEnumerator(Dummy, -1, -1);
                _rowsEnumerator = new Lut<Lut<TElement>>.ReverseLutEnumerator(
                    slice._data,
                    range.FirstPoint.Row - 1,
                    range.LastPoint.Row - 1);
            }

            public ref TElement Current => ref _columnsEnumerator.Current;

            public XLSheetPoint Point => new(_rowsEnumerator.Index + 1, _columnsEnumerator.Index + 1);

            public bool MoveNext()
            {
                while (!_columnsEnumerator.MoveNext())
                {
                    if (!_rowsEnumerator.MoveNext())
                        return false;

                    _columnsEnumerator = new Lut<TElement>.ReverseLutEnumerator(
                        _rowsEnumerator.Current,
                        _range.FirstPoint.Column - 1,
                        _range.LastPoint.Column - 1);
                }
                return true;
            }


            void IEnumerator.Reset() => throw new NotSupportedException();

            XLSheetPoint IEnumerator<XLSheetPoint>.Current => Point;

            object IEnumerator.Current => Point;
            
            public void Dispose() { }
        }
    }
}
